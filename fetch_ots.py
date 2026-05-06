#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = ["msal", "requests"]
# ///
"""
PHL5 Dashboard — OTS Miss Fetcher
Runs continuously 10:29 AM–18:05 Mon–Thu via Windows Task Scheduler.

Logic: 1 minute after each cut time → capture NOT SHIPPED from Endgame
       → write to otsData[slot_index] in data.js → dashboard shows live.
"""

import json, msal, requests, sys, time, urllib3
from datetime import date, datetime, timedelta
from pathlib import Path

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

HERE      = Path(__file__).parent
DATA_JS   = HERE / "data.js"
EG_CACHE  = HERE / ".token_cache_endgame.json"
TENANT_ID = "3cbcc3d3-094d-4006-9849-0d11d61f484d"
EG_CLIENT = "777d5c3a-eb0c-43d5-b30a-798a7eb9d15e"
EG_SCOPES = [f"{EG_CLIENT}/.default"]
FC_ID     = "3124"

# (cut_hour, cut_minute, slot_index_in_otsData)
# Check window: cut_time + 1 min  →  cut_time + 2 min  (60-second capture window)
CUT_TIMES = [
    (10, 31,  0),
    (10, 38,  1),
    (12, 31,  2),
    (13,  0,  3),
    (13, 18,  4),
    (13, 38,  5),
    (13, 58,  6),
    (14,  0,  7),
    (14,  1,  8),
    (14,  2,  9),
    (16,  8, 10),
    (16, 30, 11),
    (16, 31, 12),
    (16, 32, 13),
    (16, 48, 14),
    (16, 58, 15),
    (17, 30, 16),
    (17, 38, 17),
    (17, 58, 18),
]

STOP_AT = (18, 5)   # exit loop after 6:05 PM


# ── Endgame endpoints (tried in order until one returns NOT SHIPPED) ───────────
EG_ENDPOINTS = [
    f"https://status-api.endgame-status.prod.k8s.walmart.net/v1/api/pt-status/fc-id/{FC_ID}/summary",
    f"https://status-api.endgame-status.prod.k8s.walmart.net/v1/api/pt-status/fc-id/{FC_ID}",
    f"https://status-api.endgame-status.prod.k8s.walmart.net/v3/{FC_ID}",
]


# ── Auth ───────────────────────────────────────────────────────────────────────
def get_token() -> str:
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
        print("\n" + "=" * 60)
        print("[AUTH] One-time login required:")
        print(f"   Open : {flow['verification_uri']}")
        print(f"   Code : {flow['user_code']}")
        print("=" * 60 + "\n")
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description')}")

    if cache.has_state_changed:
        EG_CACHE.write_text(cache.serialize())

    return result["access_token"]


# ── Endgame fetch ──────────────────────────────────────────────────────────────
def fetch_not_shipped(token: str) -> int:
    """Returns the current NOT SHIPPED count from Endgame, or -1 on failure."""
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    for url in EG_ENDPOINTS:
        try:
            r = requests.get(url, headers=headers, timeout=12, verify=False)
            if r.status_code != 200:
                continue
            data = r.json()

            # summary endpoint
            if "notShipped" in data:
                return int(data["notShipped"])
            if "not_shipped" in data:
                return int(data["not_shipped"])

        except Exception as exc:
            print(f"[WARN] {url} → {exc}")
            continue

    return -1


# ── data.js helpers ────────────────────────────────────────────────────────────
def read_data_js() -> dict:
    if not DATA_JS.exists():
        return {}
    txt = DATA_JS.read_text(encoding="utf-8")
    s, e = txt.find("{"), txt.rfind("}") + 1
    if s != -1 and e:
        try:
            return json.loads(txt[s:e])
        except Exception:
            pass
    return {}


def write_data_js(payload: dict) -> None:
    ts = datetime.now().isoformat(timespec="seconds")
    payload["lastUpdated"] = ts
    DATA_JS.write_text(
        f"// Auto-generated — updated by fetch_ots.py\n"
        f"// Updated: {ts}\n"
        f"window.PHL5_DATA = {json.dumps(payload, indent=2)};\n",
        encoding="utf-8",
    )


def patch_ots_slot(idx: int, not_shipped: int) -> None:
    payload = read_data_js()
    ots = list(payload.get("otsData", [0] * 20))
    if len(ots) < 20:
        ots += [0] * (20 - len(ots))
    ots[idx] = not_shipped
    payload["otsData"] = ots
    write_data_js(payload)
    print(f"[OTS]  Slot {idx:2d} captured → NOT SHIPPED = {not_shipped:,}")


def reset_ots_if_new_day(today_str: str) -> None:
    """Wipe otsData array if data.js is from a previous day."""
    payload = read_data_js()
    last = payload.get("lastUpdated", "")[:10]
    if last and last != today_str:
        payload["otsData"] = [0] * 20
        write_data_js(payload)
        print(f"[RESET] New day detected ({last} → {today_str}). otsData cleared.")


# ── Core loop ──────────────────────────────────────────────────────────────────
def in_capture_window(now: datetime, cut_h: int, cut_m: int) -> bool:
    """True if now falls in [cut+1min, cut+2min)."""
    cut_dt    = now.replace(hour=cut_h, minute=cut_m, second=0, microsecond=0)
    window_lo = cut_dt + timedelta(minutes=1)
    window_hi = cut_dt + timedelta(minutes=2)
    return window_lo <= now < window_hi


def main() -> None:
    today = date.today()
    if today.weekday() > 3:          # Fri=4, Sat=5, Sun=6
        print("Not a shift day (Mon–Thu only). Exiting.")
        return

    today_str = today.isoformat()
    print(f"[START] OTS Fetcher — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"[INFO]  {len(CUT_TIMES)} cut times loaded. Stops at 18:05.")

    reset_ots_if_new_day(today_str)

    token = get_token()
    fired: set[int] = set()   # slots already successfully captured today

    while True:
        now = datetime.now()

        # ── Stop condition ─────────────────────────────────────────────────────
        if now.hour > STOP_AT[0] or (now.hour == STOP_AT[0] and now.minute >= STOP_AT[1]):
            print(f"[STOP]  Reached 18:05 at {now.strftime('%H:%M')}. Exiting.")
            break

        # ── Check each cut time ────────────────────────────────────────────────
        for cut_h, cut_m, idx in CUT_TIMES:
            if idx in fired:
                continue
            if in_capture_window(now, cut_h, cut_m):
                label = f"{cut_h % 12 or 12}:{cut_m:02d} {'AM' if cut_h < 12 else 'PM'}"
                print(f"[FIRE]  {now.strftime('%H:%M:%S')} — capture for {label} cut (slot {idx})")
                try:
                    ns = fetch_not_shipped(token)
                    if ns >= 0:
                        patch_ots_slot(idx, ns)
                        fired.add(idx)
                    else:
                        print(f"[RETRY] fetch failed for slot {idx} — will retry in 30s")
                except Exception as exc:
                    print(f"[ERR]   slot {idx}: {exc}")

        time.sleep(30)


if __name__ == "__main__":
    main()
