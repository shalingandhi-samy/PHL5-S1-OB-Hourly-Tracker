#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "msal",
#   "requests",
# ]
# ///
"""
PHL5 Dashboard — Data Fetcher
Sources:
  • 28-Day LP Volume  → FY26 LP - PHL5.xlsm · OB 28DP · row 64 (MS Graph)
  • 28-Day LP UPH     → FY26 LP - PHL5.xlsm · OB 28DP · row 69 (MS Graph)
  • Total Backlog     → Endgame pt-status API · FC 3124
                        Backlog = Not Shipped − Loaded − Diverted − Ship Label Applied

Runs daily at 7:30 AM Mon–Thu via Windows Task Scheduler.
"""

import json
import msal
import requests
import urllib3
from datetime import date, datetime
from pathlib import Path

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── Paths ─────────────────────────────────────────────────────────────────────
HERE         = Path(__file__).parent
DATA_JS      = HERE / "data.js"
GRAPH_CACHE  = HERE / ".token_cache_graph.json"
EG_CACHE     = HERE / ".token_cache_endgame.json"

# ── Walmart Tenant ────────────────────────────────────────────────────────────
TENANT_ID    = "3cbcc3d3-094d-4006-9849-0d11d61f484d"

# ── MS Graph (SharePoint / Excel) ─────────────────────────────────────────────
GRAPH_CLIENT = "14d82eec-204b-4c2f-b7e8-296a70dab67e"   # MS Graph Explorer
GRAPH_SCOPES = ["https://graph.microsoft.com/Files.Read.All",
                "https://graph.microsoft.com/Sites.Read.All"]
SITE_HOST    = "teams.wal-mart.com"
SITE_PATH    = "/sites/FulfillmentProductManagementAnalytics"
FILE_PATH    = "/Shared Documents/28 Day Plan/FY26 LP - PHL5.xlsm"
SHEET_NAME   = "OB 28DP"
DATE_ROW     = 7
VOL_ROW      = 64   # Processed Units Forecast
UPH_ROW      = 69   # Throughput Forecast

# ── Endgame (Backlog) ─────────────────────────────────────────────────────────
EG_CLIENT    = "777d5c3a-eb0c-43d5-b30a-798a7eb9d15e"   # Endgame client
EG_SCOPES    = [f"{EG_CLIENT}/.default"]
FC_ID        = "3124"   # PHL5

# Endpoint candidates — tried in order; first 200 response wins
EG_ENDPOINTS = [
    f"https://flow-management-api.endgame-orderpooling.prod.k8s.walmart.net/v1/{FC_ID}/flow-management/shift/recommendations",
    f"https://status-api.endgame-status.prod.k8s.walmart.net/v3/{FC_ID}",
    f"https://status-api.endgame-status.prod.k8s.walmart.net/v1/api/pt-status/fc-id/{FC_ID}/summary",
    f"https://status-api.endgame-status.prod.k8s.walmart.net/v1/api/pt-status/fc-id/{FC_ID}",
    f"https://api.order-planning.prod.k8s.walmart.net/v1/api/pt-status/fc-id/{FC_ID}/blitz-stats",
]


# ══════════════════════════════════════════════════════════════════════════════
# AUTH HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _msal_app(client_id: str, cache_file: Path) -> tuple:
    cache = msal.SerializableTokenCache()
    if cache_file.exists():
        cache.deserialize(cache_file.read_text())
    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )
    return app, cache


def _get_token(client_id: str, scopes: list, cache_file: Path, label: str) -> str:
    app, cache = _msal_app(client_id, cache_file)
    accounts = app.get_accounts()

    result = None
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result or "access_token" not in result:
        flow = app.initiate_device_flow(scopes=scopes)
        print(f"\n{'='*60}")
        print(f"🔐 [{label}] Auth required — open this URL in your browser:")
        print(f"   {flow['verification_uri']}")
        print(f"   Enter code: {flow['user_code']}")
        print(f"{'='*60}\n")
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"[{label}] Auth failed: {result.get('error_description', result)}")

    if cache.has_state_changed:
        cache_file.write_text(cache.serialize())

    return result["access_token"]


def graph_token() -> str:
    return _get_token(GRAPH_CLIENT, GRAPH_SCOPES, GRAPH_CACHE, "MS Graph")


def endgame_token() -> str:
    return _get_token(EG_CLIENT, EG_SCOPES, EG_CACHE, "Endgame")


# ══════════════════════════════════════════════════════════════════════════════
# MS GRAPH — SharePoint Excel
# ══════════════════════════════════════════════════════════════════════════════

def graph_get(token: str, url: str) -> dict:
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    r.raise_for_status()
    return r.json()


def get_workbook_ids(token: str) -> tuple[str, str]:
    site = graph_get(token, f"https://graph.microsoft.com/v1.0/sites/{SITE_HOST}:{SITE_PATH}")
    site_id = site["id"]
    encoded = FILE_PATH.replace(" ", "%20")
    item = graph_get(token, f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{encoded}")
    return site_id, item["id"]


def read_range(token: str, site_id: str, item_id: str, address: str) -> list:
    sheet = SHEET_NAME.replace(" ", "%20")
    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}"
           f"/workbook/worksheets/{sheet}/range(address='{address}')/values")
    return graph_get(token, url).get("values", [[]])


def today_col_letter(token: str, site_id: str, item_id: str) -> str:
    label = f"{date.today().day}-{date.today().strftime('%b')}"
    print(f"🔍 Scanning row {DATE_ROW} for '{label}'...")
    row = read_range(token, site_id, item_id, f"F{DATE_ROW}:HZ{DATE_ROW}")
    cells = row[0] if row else []
    for i, v in enumerate(cells):
        if str(v).strip() == label:
            n = 6 + i  # F=6 (A=1)
            letter = _col_letter(n)
            print(f"✅ Found at column {letter}")
            return letter
    raise ValueError(f"Date '{label}' not found in row {DATE_ROW}")


def _col_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


# ══════════════════════════════════════════════════════════════════════════════
# ENDGAME — Backlog
# ══════════════════════════════════════════════════════════════════════════════

def fetch_backlog(token: str) -> tuple[int, str]:
    """
    Try each Endgame endpoint until one returns usable data.
    Returns (backlog_units, endpoint_used).
    Formula: Backlog = Not Shipped − Loaded − Diverted − Ship Label Applied
    """
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    for url in EG_ENDPOINTS:
        try:
            r = requests.get(url, headers=headers, timeout=12, verify=False)
            if r.status_code != 200:
                print(f"  ↳ {url.split('/')[-2:]}: HTTP {r.status_code}")
                continue

            data = r.json()
            print(f"  ✅ Got data from: {url}")

            # ── Flow Management shift/recommendations schema ────────────────
            # { backlog: [{ processPath, units, brands }], ... }
            if "backlog" in data and isinstance(data["backlog"], list):
                total = sum(b.get("units", 0) for b in data["backlog"])
                print(f"  📦 backlog (flow-mgmt sum): {total:,}")
                return total, url

            # ── pt-status summary schema ───────────────────────────────────
            # { notShipped, loaded, diverted, shippingLabelApplied, ... }
            if "notShipped" in data or "not_shipped" in data:
                not_shipped   = _dig(data, "notShipped",          "not_shipped",          0)
                loaded        = _dig(data, "loaded",                                       0)
                diverted      = _dig(data, "diverted",                                     0)
                ship_label    = _dig(data, "shippingLabelApplied", "ship_label_applied",   0)
                backlog = max(0, not_shipped - loaded - diverted - ship_label)
                print(f"  📦 not_shipped={not_shipped:,} loaded={loaded:,} "
                      f"diverted={diverted:,} shipLabel={ship_label:,} → backlog={backlog:,}")
                return backlog, url

            # ── GetShiftSetup v3 ───────────────────────────────────────────
            if "processPathList" in data:
                total = sum(p.get("overdueUnits", 0) for p in data.get("processPathList", []))
                print(f"  📦 overdueUnits (shift setup sum): {total:,}")
                return total, url

            # Dump keys so we can learn the schema
            print(f"  ⚠️  Unknown schema. Top-level keys: {list(data.keys())[:15]}")

        except requests.exceptions.ConnectionError as e:
            print(f"  ↳ Connection error: {e}")
        except Exception as e:
            print(f"  ↳ Error: {e}")

    print("⚠️  No Endgame endpoint returned usable backlog data.")
    return 0, "—"


def _dig(d: dict, *keys, default=0):
    """Return first matching key's value from dict, or default."""
    for k in keys:
        if k in d:
            v = d[k]
            # May be a dict with a 'units' or 'count' sub-field
            if isinstance(v, dict):
                return v.get("units", v.get("count", v.get("total", default)))
            if isinstance(v, (int, float)):
                return int(v)
    return default


# ══════════════════════════════════════════════════════════════════════════════
# OUTPUT
# ══════════════════════════════════════════════════════════════════════════════

def write_data_js(lp_volume=0, lp_uph=0, backlog=0,
                  cell_vol="—", cell_uph="—", skipped=False):
    ts = datetime.now().isoformat(timespec="seconds")
    payload = {
        "lpVolume":    lp_volume or 0,
        "lpUPH":       lp_uph    or 0,
        "backlog":     backlog   or 0,
        "cellVol":     cell_vol,
        "cellUPH":     cell_uph,
        "lastUpdated": ts,
        "skipped":     skipped,
    }
    js = (f"// Auto-generated by fetch_lp_data.py\n"
          f"// Updated: {ts}\n"
          f"window.PHL5_DATA = {json.dumps(payload, indent=2)};\n")
    DATA_JS.write_text(js, encoding="utf-8")
    print(f"💾 Written → {DATA_JS}")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    today = date.today()

    if today.weekday() > 3:  # Fri=4, Sat=5, Sun=6
        print(f"📅 Today is {today.strftime('%A')} — S1 runs Mon–Thu. Skipping.")
        write_data_js(skipped=True)
        return

    print(f"🚀 PHL5 Data Fetcher — {today.strftime('%A, %B %d, %Y')}")

    lp_volume, lp_uph, backlog = 0, 0, 0
    cell_vol, cell_uph = "—", "—"

    # ── 1. MS Graph: LP Volume + LP UPH ─────────────────────────────────────
    try:
        g_token           = graph_token()
        site_id, item_id  = get_workbook_ids(g_token)
        col               = today_col_letter(g_token, site_id, item_id)

        cell_vol = f"{col}{VOL_ROW}"
        raw      = read_range(g_token, site_id, item_id, cell_vol)
        lp_volume = int(float(str(raw[0][0]).replace(",", ""))) if raw and raw[0] else 0
        print(f"✅ LP Volume ({cell_vol}): {lp_volume:,}")

        cell_uph = f"{col}{UPH_ROW}"
        raw      = read_range(g_token, site_id, item_id, cell_uph)
        lp_uph   = round(float(str(raw[0][0]).replace(",", "")), 1) if raw and raw[0] else 0
        print(f"✅ LP UPH   ({cell_uph}): {lp_uph}")
    except Exception as e:
        print(f"❌ Graph fetch failed: {e}")

    # ── 2. Endgame: Backlog ─────────────────────────────────────────────────
    try:
        print("\n📡 Fetching backlog from Endgame...")
        eg_token      = endgame_token()
        backlog, _url = fetch_backlog(eg_token)
        print(f"✅ Backlog: {backlog:,}")
    except Exception as e:
        print(f"❌ Endgame fetch failed: {e}")

    # ── 3. Write output ─────────────────────────────────────────────────────
    write_data_js(lp_volume=lp_volume, lp_uph=lp_uph, backlog=backlog,
                  cell_vol=cell_vol, cell_uph=cell_uph)


if __name__ == "__main__":
    main()
