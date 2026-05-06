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
  • actualVol, asrsUPH, hoursWorked                          ← DRAX Stationary Picking
  • actualBag                                                ← DRAX Bagging - Manual
  • boxUPH                                                   ← DRAX Box Finishing
  • capacity (1P / 3P/WFS / SAMS)                           ← Gscope Distributor Capacity
  • otsData (20-slot miss counts per cut time)               ← Endgame /cutoffs

Auth:
  • Endgame  — MSAL device-code (cached in .token_cache_endgame.json)
  • DRAX     — browser session cookie  → DRAX_COOKIE in .env
  • Gscope   — gateway_token / cookie  → GSCOPE_COOKIE or GSCOPE_TOKEN in .env
"""

import json
import logging
import msal
import os
import re
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
DRAX_BASE_URL    = os.getenv("DRAX_BASE_URL",        "https://drax.walmart.com")
DRAX_COOKIE      = os.getenv("DRAX_COOKIE",          "")

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

# Cut times hardcoded in index.html — order MUST match the 20-slot otsData array.
# Empty string at index 19 = spare slot (no cut time assigned).
CUT_TIMES = [
    "10:31 AM", "10:38 AM", "12:31 PM",  "1:00 PM",  "1:18 PM",
     "1:38 PM",  "1:58 PM",  "2:00 PM",  "2:01 PM",  "2:02 PM",
     "4:08 PM",  "4:30 PM",  "4:31 PM",  "4:32 PM",  "4:48 PM",
     "4:58 PM",  "5:30 PM",  "5:38 PM",  "5:58 PM",  "",
]

# Shift: 7:30 AM–6:00 PM → 11 hourly slots.
# Wall-clock hour (0-23) for each slot — matches DRAX column index within a day.
SHIFT_HOURS = [7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]
HOUR_TO_IDX = {h: i for i, h in enumerate(SHIFT_HOURS)}

# DRAX HTML scrape — one page pass, three departments.
_COLS_PER_DAY = 24
_DEPT_SCRAPE: list[tuple[str, dict[str, str]]] = [
    ("Stationary Picking", {"units": "volume", "hours": "hours_worked", "uph": "asrs_uph"}),
    ("Bagging - Manual",   {"units": "bagging"}),
    ("Box Finishing",      {"uph":   "box_finish_uph"}),
]
_METRIC_PAT: dict[str, re.Pattern] = {
    m: re.compile(rf"matrixmetric\s*=\s*'{m}'>([ \d,.]+)<", re.IGNORECASE)
    for m in ("units", "hours", "uph")
}


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
# DRAX
# ══════════════════════════════════════════════════════════════════════════════

def _days_since_saturday(d: date | None = None) -> int:
    """Walmart fiscal week starts Saturday. Returns 0=Sat, 1=Sun, 2=Mon …"""
    d = d or date.today()
    return (d.weekday() - 5) % 7


def _drax_headers() -> dict:
    return {
        "Accept": "text/html,application/xhtml+xml,*/*;q=0.9",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
        ),
        "Cookie":   DRAX_COOKIE,
        "Referer":  f"{DRAX_BASE_URL}/",
    }


def _parse_drax_html(html: str, d: date) -> dict | None:
    day_offset = _days_since_saturday(d)
    col_start  = day_offset * _COLS_PER_DAY
    col_end    = col_start + _COLS_PER_DAY

    raw: dict[str, dict[str, list[str]]] = {}
    needed = {dept for dept, _ in _DEPT_SCRAPE}

    for row in html.split("<tr"):
        for dept_name, metric_map in _DEPT_SCRAPE:
            if dept_name in raw:
                continue
            if f"department='{dept_name}'" not in row:
                continue
            first_metric = next(iter(metric_map))
            if not _METRIC_PAT[first_metric].search(row):
                continue
            raw[dept_name] = {m: _METRIC_PAT[m].findall(row) for m in metric_map}
        if raw.keys() >= needed:
            break

    sp = "Stationary Picking"
    if sp not in raw:
        log.warning("DRAX: Stationary Picking row not found — cookie may be stale")
        return None

    sp_units = raw[sp].get("units", [])
    if not sp_units[col_start:col_end]:                       # fallback to prev day
        col_start = max(0, col_start - _COLS_PER_DAY)
        col_end   = col_start + _COLS_PER_DAY

    n_slots = len(sp_units[col_start:col_end])
    # hour = i means wall-clock hour within the day (0=midnight … 23=11pm).
    # col_start is a weekly-matrix offset, NOT the hour — do not add it here.
    slots: list[dict] = [{"hour": i} for i in range(n_slots)]

    for dept_name, metric_map in _DEPT_SCRAPE:
        if dept_name not in raw:
            continue
        for html_metric, slot_key in metric_map.items():
            today_vals = raw[dept_name].get(html_metric, [])[col_start:col_end]
            is_int = slot_key in ("volume", "bagging")
            for i, raw_str in enumerate(today_vals[:n_slots]):
                try:
                    val = (int(raw_str.replace(",", "")) if is_int
                           else float(raw_str.replace(",", "")))
                    slots[i][slot_key] = val
                except ValueError:
                    pass

    all_zero = all(s.get("volume", 0) == 0 for s in slots)
    if all_zero:
        log.warning(
            "DRAX: today's slice [%d:%d] is all-zero — "
            "data not yet published (Drax batches after shift close). "
            "Returning None to preserve existing data.js values.",
            col_start, col_end,
        )
        return None

    log.info("DRAX: parsed %d slots for %s", n_slots, d)
    return {"hours": slots}


def fetch_drax(target_date: date | None = None) -> dict:
    """Fetch DRAX building_overview; return shift arrays or zeros on failure."""
    empty_arrays = {
        "actualVol":   [0]   * len(SHIFT_HOURS),
        "asrsUPH":     [0.0] * len(SHIFT_HOURS),
        "boxUPH":      [0.0] * len(SHIFT_HOURS),
        "actualBag":   [0]   * len(SHIFT_HOURS),
        "hoursWorked": [0.0] * len(SHIFT_HOURS),
    }

    if not DRAX_COOKIE:
        log.warning("DRAX_COOKIE not set in .env — skipping DRAX fetch")
        return empty_arrays

    d = target_date or date.today()
    date_str = d.strftime("%Y-%m-%d")
    url = (
        f"{DRAX_BASE_URL}/building_overview/"
        f"?date_hour_after={date_str}+07:00"
        f"&date_hour_before={date_str}+23:59"
        f"&area=Outbound"
    )

    try:
        r = requests.get(url, headers=_drax_headers(), timeout=30, verify=False)
    except Exception as e:
        log.warning("DRAX connection error: %s", e)
        return empty_arrays

    if r.status_code in (401, 403):
        log.warning("DRAX auth failed (HTTP %s) — cookie expired", r.status_code)
        return empty_arrays
    if r.status_code != 200:
        log.warning("DRAX returned HTTP %s", r.status_code)
        return empty_arrays

    parsed = _parse_drax_html(r.text, d)
    if not parsed:
        return empty_arrays

    # Map wall-clock hour slots → 11-element shift arrays
    result = {k: list(v) for k, v in empty_arrays.items()}
    for slot in parsed["hours"]:
        wall_hour = slot.get("hour", -1)
        idx = HOUR_TO_IDX.get(wall_hour)
        if idx is None:
            continue
        result["actualVol"][idx]   = slot.get("volume",       0)
        result["asrsUPH"][idx]     = slot.get("asrs_uph",     0.0)
        result["boxUPH"][idx]      = slot.get("box_finish_uph",0.0)
        result["actualBag"][idx]   = slot.get("bagging",      0)
        result["hoursWorked"][idx] = slot.get("hours_worked", 0.0)

    return result


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

    # ── 2. DRAX — Hourly Volume / UPH / Bagging / Hours ────────────────────
    print("📡 DRAX hourly metrics...")
    drax = fetch_drax()
    updates.update(drax)
    total_vol = sum(drax["actualVol"])
    print(f"   {'✅' if total_vol else '⚠️ '} actualVol total={total_vol:,} units  "
          f"| DRAX_COOKIE={'set' if DRAX_COOKIE else 'MISSING — set in .env'}")

    # ── 3. Gscope — OB Capacity ─────────────────────────────────────────────
    print("📡 Gscope OB capacity...")
    capacity = fetch_gscope()
    if capacity:
        updates["capacity"] = capacity
        for c in capacity:
            print(f"   ✅ {c['type']:<7} planned={c['planned']:>6,}  "
                  f"consumed={c['consumed']:>6,}  avail={c['available']:>6,}")
    else:
        print("   ⚠️  Gscope skipped — set GSCOPE_COOKIE or GSCOPE_TOKEN in .env")

    # ── 4. Write ─────────────────────────────────────────────────────────────

    patch_data_js(**updates)
    print(f"\n✅ Done — {updates.get('lastUpdated', '?')}\n")


if __name__ == "__main__":
    main()
