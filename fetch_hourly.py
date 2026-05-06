#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "msal",
#   "requests",
#   "python-dotenv",
#   "urllib3",
# ]
# ///
"""
PHL5 OB Hourly Fetch — runs at :15 of each hour, Mon–Thu, 8:15 AM–5:15 PM.

Patches data.js with:
  • Backlog, blNotShipped, blLoaded, blDiverted, blShipLabel  ← Endgame FC 3124
  • capacity (1P / 3P/WFS / SAMS)                           ← Gscope Distributor Capacity
  • otsData (20-slot miss counts per cut time)               ← Endgame /cutoffs

NOTE: DRAX fields (actualVol, asrsUPH, hoursWorked, actualBag, boxUPH) are written
      by the qa-kitten scheduler task (PHL5_Backlog_QA) — one browser visit, no cookie needed.

Auth:
  • Endgame  — MSAL device-code (cached in .token_cache_endgame.json)
  • Gscope   — gateway_token / cookie  → GSCOPE_COOKIE or GSCOPE_TOKEN in .env
"""

import json
import logging
import msal
import os
import requests
import sys
import urllib3
from datetime import date, datetime
from pathlib import Path

# Load .env if present (python-dotenv)
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent / ".env")
except ImportError:
    pass

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

logging.basicConfig(level=logging.INFO, format="%(levelname)s  %(message)s")
log = logging.getLogger(__name__)

# ── Paths ─────────────────────────────────────────────────────────────────────
HERE     = Path(__file__).parent
DATA_JS  = HERE / "data.js"
EG_CACHE = HERE / ".token_cache_endgame.json"

# ── Config ────────────────────────────────────────────────────────────────────
GSCOPE_BASE_URL  = os.getenv("GSCOPE_BASE_URL",      "https://gscope.walmart.com")
GSCOPE_COOKIE    = os.getenv("GSCOPE_COOKIE",        "")
GSCOPE_TOKEN     = os.getenv("GSCOPE_TOKEN",         "")
GSCOPE_DIST_ID   = os.getenv("GSCOPE_DISTRIBUTOR_ID","31341")   # PHL5
GSCOPE_POOL_1P   = int(os.getenv("GSCOPE_POOL_1P",  "1"))
GSCOPE_POOL_3P   = int(os.getenv("GSCOPE_POOL_3P",  "4"))
GSCOPE_POOL_SAMS = int(os.getenv("GSCOPE_POOL_SAMS","27"))

TENANT_ID  = "3cbcc3d3-094d-4006-9849-0d11d61f484d"
EG_CLIENT  = "777d5c3a-eb0c-43d5-b30a-798a7eb9d15e"
EG_SCOPES  = [f"{EG_CLIENT}/.default"]
FC_ID      = "3124"

EG_ENDPOINTS = [
    f"https://flow-management-api.endgame-orderpooling.prod.k8s.walmart.net/v1/{FC_ID}/flow-management/shift/recommendations",
    f"https://status-api.endgame-status.prod.k8s.walmart.net/v3/{FC_ID}",
    f"https://status-api.endgame-status.prod.k8s.walmart.net/v1/api/pt-status/fc-id/{FC_ID}/summary",
    f"https://status-api.endgame-status.prod.k8s.walmart.net/v1/api/pt-status/fc-id/{FC_ID}",
]

EG_STATUS_API = "https://status-api.endgame-status.prod.k8s.walmart.net"
FCAP_BASE     = "https://gscope.walmart.com/api/gateway/fcap"

# Cut times hardcoded in index.html — DRAX arrays are owned by qa-kitten, not this script. — order MUST match the 20-slot otsData array.
# Empty string at index 19 = spare slot (no cut time assigned).
CUT_TIMES = [
    "10:31 AM", "10:38 AM", "12:31 PM",  "1:00 PM",  "1:18 PM",
     "1:38 PM",  "1:58 PM",  "2:00 PM",  "2:01 PM",  "2:02 PM",
     "4:08 PM",  "4:30 PM",  "4:31 PM",  "4:32 PM",  "4:48 PM",
     "4:58 PM",  "5:30 PM",  "5:38 PM",  "5:58 PM",  "",
]



# ══════════════════════════════════════════════════════════════════════════════
# ENDGAME
# ══════════════════════════════════════════════════════════════════════════════

def _endgame_token() -> str:
    cache = msal.SerializableTokenCache()
    if EG_CACHE.exists():
        cache.deserialize(EG_CACHE.read_text())
    app = msal.PublicClientApplication(
        EG_CLIENT,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )
    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(EG_SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        flow = app.initiate_device_flow(scopes=EG_SCOPES)
        print(f"\n{'='*60}")
        print("[AUTH] Endgame login required:")
        print(f"   Open: {flow['verification_uri']}")
        print(f"   Code: {flow['user_code']}")
        print(f"{'='*60}\n")
        result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Endgame auth failed: {result.get('error_description')}")
    if cache.has_state_changed:
        EG_CACHE.write_text(cache.serialize())
    return result["access_token"]


def fetch_endgame() -> dict:
    """Returns backlog + component fields. Falls back to zeros on failure."""
    defaults = dict(backlog=0, blNotShipped=0, blLoaded=0, blDiverted=0, blShipLabel=0)
    try:
        token = _endgame_token()
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        for url in EG_ENDPOINTS:
            try:
                r = requests.get(url, headers=headers, timeout=12, verify=False)
                if r.status_code != 200:
                    continue
                data = r.json()

                if "backlog" in data and isinstance(data["backlog"], list):
                    total = sum(b.get("units", 0) for b in data["backlog"])
                    return dict(backlog=total, blNotShipped=total,
                                blLoaded=0, blDiverted=0, blShipLabel=0)

                if any(k in data for k in ("notShipped", "not_shipped")):
                    ns  = int(data.get("notShipped",          data.get("not_shipped", 0)))
                    ld  = int(data.get("loaded",              0))
                    div = int(data.get("diverted",            0))
                    sla = int(data.get("shippingLabelApplied",data.get("ship_label_applied", 0)))
                    return dict(backlog=max(0, ns - ld - div - sla),
                                blNotShipped=ns, blLoaded=ld,
                                blDiverted=div, blShipLabel=sla)

                if "processPathList" in data:
                    total = sum(p.get("overdueUnits", 0) for p in data["processPathList"])
                    return dict(backlog=total, blNotShipped=total,
                                blLoaded=0, blDiverted=0, blShipLabel=0)

            except Exception as e:
                log.debug("Endgame endpoint error: %s", e)
        log.warning("Endgame: no usable response from any endpoint")
    except Exception as e:
        log.error("Endgame fetch failed: %s", e)
    return defaults


# Status key variants Endgame uses for "Not Shipped" — checked case-insensitively.
_NOT_SHIPPED_KEYS = {"NOT_SHIPPED", "NOTSHIPPED", "NOT SHIPPED"}


def fetch_ots(token: str, current_ots: list[int]) -> list[int]:
    """Fetch OTS miss (= NOT SHIPPED count) per cut time from Endgame /cutoffs.

    Only past cut times are updated — future ones carry forward unchanged.
    Returns a 20-element int array aligned to CUT_TIMES.
    """
    result = list(current_ots)          # carry forward existing values
    now    = datetime.now()
    url    = f"{EG_STATUS_API}/{FC_ID}/cutoffs"

    try:
        r = requests.get(
            url,
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
            params={"date": date.today().isoformat()},
            timeout=15,
            verify=False,
        )
        if r.status_code != 200:
            log.warning("Endgame /cutoffs returned HTTP %s", r.status_code)
            return result

        # Build {normalised_time: not_shipped_count} — NOT SHIPPED column only
        not_shipped: dict[str, int] = {}
        for _group, group in r.json().get("cutOffGroups", {}).items():
            for time_str, statuses in group.get("cutOffTimes", {}).items():
                for status_key, bucket in statuses.items():
                    norm_key = status_key.upper().replace("-", "_").replace(" ", "_")
                    if norm_key not in _NOT_SHIPPED_KEYS:
                        continue
                    if isinstance(bucket, dict):
                        cnt = bucket.get("count", {})
                        if isinstance(cnt, dict):
                            val = int(cnt.get("pickTickets", 0) or 0)
                            key = _normalise_time(time_str)
                            not_shipped[key] = not_shipped.get(key, 0) + val

        # Update only cut times that have already passed
        updated = 0
        for i, ct in enumerate(CUT_TIMES):
            if not ct:
                continue
            try:
                ct_dt = datetime.strptime(ct, "%I:%M %p").replace(
                    year=now.year, month=now.month, day=now.day
                )
            except ValueError:
                continue
            if now >= ct_dt:
                result[i] = not_shipped.get(_normalise_time(ct), 0)
                updated += 1

        log.info("OTS: updated %d past cut times, total not-shipped=%d", updated, sum(result))
        return result

    except Exception as e:
        log.error("Endgame /cutoffs fetch failed: %s", e)
        return result


def _normalise_time(t: str) -> str:
    """Parse a time string and reformat as '12:34 PM' for reliable comparison."""
    if not t.strip():
        return ""
    for fmt in ("%I:%M %p", "%H:%M", "%I:%M%p"):
        try:
            return datetime.strptime(t.strip(), fmt).strftime("%I:%M %p").lstrip("0")
        except ValueError:
            continue
    return t.strip()  # fallback — keep as-is


# ══════════════════════════════════════════════════════════════════════════════
# GSCOPE
# ══════════════════════════════════════════════════════════════════════════════

def _gscope_headers() -> dict:
    h = {
        "Accept": "application/json",
        "Referer": f"{GSCOPE_BASE_URL}/mfe/distributorCapacity",
        "Origin": "https://gscope.walmart.com",
        "X-Requested-With": "XMLHttpRequest",
    }
    bearer = GSCOPE_TOKEN.replace("Bearer ", "").strip() if GSCOPE_TOKEN else None
    if GSCOPE_COOKIE:
        h["Cookie"] = GSCOPE_COOKIE
        if not bearer:
            for part in GSCOPE_COOKIE.split(";"):
                p = part.strip()
                if p.startswith("gateway_token="):
                    bearer = p[len("gateway_token="):]
    if bearer:
        h["Authorization"] = f"Bearer {bearer}"
    return h


def _fetch_pool(distributor_id: str, pool_id: int) -> tuple[int, int]:
    """Returns (planned, consumed) for one Gscope capacity pool."""
    url = f"{FCAP_BASE}/getFCDetail?distributorId={distributor_id}&poolId={pool_id}"
    try:
        r = requests.get(url, headers=_gscope_headers(), timeout=15, verify=False)
        if r.status_code != 200:
            log.debug("Gscope pool %d: HTTP %d", pool_id, r.status_code)
            return 0, 0

        payload = r.json().get("payload", r.json())

        planned = payload.get("plannedCapacity")
        if planned is None:
            for resp in payload.get("capacityPathResponses", []):
                for node in resp.get("nodeCapacityDetails", []):
                    try:
                        planned = (node["nodeConfiguration"]["capacityMeasure"]
                                   ["absoluteQuantity"]["value"])
                        break
                    except (KeyError, TypeError):
                        continue
                if planned is not None:
                    break

        available = (payload.get("availableCapacity")
                     or payload.get("available_capacity"))

        planned   = int(str(planned  ).replace(",", "")) if planned   is not None else 0
        available = int(str(available).replace(",", "")) if available is not None else 0
        consumed  = max(0, planned - available)
        return planned, consumed

    except Exception as e:
        log.debug("Gscope pool %d error: %s", pool_id, e)
        return 0, 0


def fetch_gscope() -> list[dict]:
    """Returns capacity rows for 1P, 3P/WFS, SAMS, or carry-forward zeros."""
    pools = [
        ("1P",     GSCOPE_POOL_1P,   "#0053e2"),
        ("3P/WFS", GSCOPE_POOL_3P,   "#ffc220"),
        ("SAMS",   GSCOPE_POOL_SAMS, "#2a8703"),
    ]
    if not (GSCOPE_COOKIE or GSCOPE_TOKEN):
        log.warning("GSCOPE_COOKIE / GSCOPE_TOKEN not set in .env — skipping Gscope")
        return []

    rows = []
    for network, pool_id, color in pools:
        planned, consumed = _fetch_pool(GSCOPE_DIST_ID, pool_id)
        available = max(0, planned - consumed)
        rows.append({"type": network, "pool": pool_id,
                     "planned": planned, "consumed": consumed,
                     "available": available, "color": color})
        log.info("Gscope %-7s pool %2d → planned=%6d consumed=%6d avail=%6d",
                 network, pool_id, planned, consumed, available)
    return rows


# ══════════════════════════════════════════════════════════════════════════════
# DATA.JS PATCHER
# ══════════════════════════════════════════════════════════════════════════════

def patch_data_js(**updates: object) -> None:
    current: dict = {}
    if DATA_JS.exists():
        txt = DATA_JS.read_text(encoding="utf-8")
        start = txt.find("{")
        end   = txt.rfind("}") + 1
        if start != -1 and end:
            try:
                current = json.loads(txt[start:end])
            except Exception:
                pass

    current.update(updates)
    current["lastUpdated"] = datetime.now().isoformat(timespec="seconds")
    ts = current["lastUpdated"]
    DATA_JS.write_text(
        f"// Auto-generated by fetch_hourly.py\n"
        f"// Updated: {ts}\n"
        f"window.PHL5_DATA = {json.dumps(current, indent=2)};\n",
        encoding="utf-8",
    )
    log.info("data.js patched → %s", DATA_JS.name)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    today = date.today()
    if today.weekday() > 3:  # Fri=4 … Sun=6
        log.info("Not a shift day (%s) — skipping.", today.strftime("%A"))
        return

    print(f"\n🚀 PHL5 Hourly Fetch — {datetime.now().strftime('%H:%M:%S  %a %b %d %Y')}\n")
    updates: dict = {}

    # ── 1. Endgame — Backlog + OTS ─────────────────────────────────────────
    print("📡 Endgame backlog + OTS...")
    eg = fetch_endgame()
    updates.update(eg)
    print(f"   ✅ Backlog={eg['backlog']:,}  "
          f"(NS={eg['blNotShipped']:,}  L={eg['blLoaded']:,}  "
          f"D={eg['blDiverted']:,}  SLA={eg['blShipLabel']:,})")

    # OTS — reuses cached Endgame token, carries forward existing values for future cut times
    try:
        current_txt = DATA_JS.read_text(encoding="utf-8") if DATA_JS.exists() else "{}"
        start = current_txt.find("{")
        end   = current_txt.rfind("}") + 1
        current_ots = [0] * len(CUT_TIMES)
        if start != -1 and end:
            try:
                current_ots = json.loads(current_txt[start:end]).get("otsData", current_ots)
            except Exception:
                pass

        eg_token = _endgame_token()
        ots = fetch_ots(eg_token, current_ots)
        updates["otsData"] = ots
        past_ct  = sum(1 for ct in CUT_TIMES if ct and
                       datetime.strptime(ct, "%I:%M %p").replace(
                           year=datetime.now().year, month=datetime.now().month,
                           day=datetime.now().day) <= datetime.now())
        print(f"   ✅ OTS not-shipped total={sum(ots):,}  "
              f"({past_ct} past cut times checked)")
    except Exception as e:
        log.error("OTS fetch failed: %s", e)

    # ── 2. Gscope — OB Capacity ─────────────────────────────────────────────
    print("📡 Gscope OB capacity...")
    capacity = fetch_gscope()
    if capacity:
        updates["capacity"] = capacity
        for c in capacity:
            print(f"   ✅ {c['type']:<7} planned={c['planned']:>6,}  "
                  f"consumed={c['consumed']:>6,}  avail={c['available']:>6,}")
    else:
        print("   ⚠️  Gscope skipped — set GSCOPE_COOKIE or GSCOPE_TOKEN in .env")

    # ── 3. Write ─────────────────────────────────────────────────────────────

    patch_data_js(**updates)
    print(f"\n✅ Done — {updates.get('lastUpdated', '?')}\n")


if __name__ == "__main__":
    main()
