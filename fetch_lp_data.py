#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "msal",
#   "requests",
# ]
# ///
"""
PHL5 Dashboard — LP Data Fetcher
Reads OB 28DP row 64 (Processed Units Forecast) for today's date
from the FY26 LP - PHL5.xlsm SharePoint file and writes data.js.

Runs daily at 7:30 AM Mon–Thu via Code Puppy Scheduler.
"""

import json
import sys
import os
import msal
import requests
from datetime import date, datetime
from pathlib import Path

# ── Constants ────────────────────────────────────────────────────────────────
TENANT_ID   = "3cbcc3d3-094d-4006-9849-0d11d61f484d"   # Walmart tenant
CLIENT_ID   = "14d82eec-204b-4c2f-b7e8-296a70dab67e"   # MS Graph Explorer
SCOPES      = ["https://graph.microsoft.com/Files.Read.All",
                "https://graph.microsoft.com/Sites.Read.All"]

SITE_HOST   = "teams.wal-mart.com"
SITE_PATH   = "/sites/FulfillmentProductManagementAnalytics"
FILE_PATH   = "/Shared Documents/28 Day Plan/FY26 LP - PHL5.xlsm"
SHEET_NAME  = "OB 28DP"
DATE_ROW    = 7     # Row containing daily dates (d-mmm format)
FORECAST_ROW = 64  # Row containing Processed Units Forecast
UPH_ROW      = 69  # Row containing Throughput Forecast (LP UPH)

DATA_JS     = Path(__file__).parent / "data.js"
TOKEN_CACHE = Path(__file__).parent / ".token_cache.json"


# ── Token cache (persist between runs) ───────────────────────────────────────
def build_msal_app() -> msal.PublicClientApplication:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE.exists():
        cache.deserialize(TOKEN_CACHE.read_text())
    app = msal.PublicClientApplication(CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}", token_cache=cache)
    return app, cache


def get_token() -> str:
    app, cache = build_msal_app()
    accounts = app.get_accounts()

    # Silent auth (uses cached token)
    result = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    # Device code flow (first-time / expired)
    if not result or "access_token" not in result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        print("\n" + "="*60)
        print("🔐 Authentication required — open this URL in your browser:")
        print(f"   {flow['verification_uri']}")
        print(f"   Enter code: {flow['user_code']}")
        print("="*60 + "\n")
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")

    # Persist cache
    if cache.has_state_changed:
        TOKEN_CACHE.write_text(cache.serialize())

    return result["access_token"]


# ── Graph API helpers ─────────────────────────────────────────────────────────
def graph_get(token: str, url: str) -> dict:
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    r.raise_for_status()
    return r.json()


def get_site_id(token: str) -> str:
    data = graph_get(token, f"https://graph.microsoft.com/v1.0/sites/{SITE_HOST}:{SITE_PATH}")
    return data["id"]


def get_drive_item_id(token: str, site_id: str) -> str:
    """Locate the .xlsm file in the drive by path."""
    encoded = FILE_PATH.replace(" ", "%20")
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{encoded}"
    data = graph_get(token, url)
    return data["id"], site_id


def read_range(token: str, site_id: str, item_id: str, address: str) -> list:
    """Read a cell range from the OB 28DP sheet. Returns list of rows (list of lists)."""
    sheet = SHEET_NAME.replace(" ", "%20")
    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}"
           f"/workbook/worksheets/{sheet}/range(address='{address}')/values")
    data = graph_get(token, url)
    return data.get("values", [[]])


# ── Date matching ─────────────────────────────────────────────────────────────
def today_label() -> str:
    """Return today's date in Excel d-mmm format (e.g. '5-May')."""
    t = date.today()
    return f"{t.day}-{t.strftime('%b')}"


def find_today_col(token: str, site_id: str, item_id: str) -> str:
    """
    Scan row 7 (date row) across columns F–HZ to find today's date.
    Returns the column letter (e.g. 'HA').
    """
    label = today_label()
    print(f"🔍 Looking for date label '{label}' in row {DATE_ROW}...")

    # Read a wide range — F7:HZ7
    values = read_range(token, site_id, item_id, f"F{DATE_ROW}:HZ{DATE_ROW}")
    row = values[0] if values else []

    # Column F = index 0, G = 1, ...
    col_start_num = 6  # F = column 6 (A=1)
    for i, cell_val in enumerate(row):
        if str(cell_val).strip() == label:
            col_num = col_start_num + i
            col_letter = col_num_to_letter(col_num)
            print(f"✅ Found '{label}' at column {col_letter} (index {i})")
            return col_letter

    raise ValueError(f"Date '{label}' not found in row {DATE_ROW}. Is today a scheduled date?")


def col_num_to_letter(n: int) -> str:
    """Convert 1-based column number to Excel letter (e.g. 1→A, 27→AA, 209→HA)."""
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    today = date.today()

    # Only run Mon–Thu (weekday 0–3)
    if today.weekday() > 3:
        print(f"📅 Today is {today.strftime('%A')} — S1 shift is Mon–Thu only. Skipping.")
        write_data_js(lp_volume=None, lp_uph=None, skipped=True)
        return

    print(f"🚀 PHL5 LP Data Fetcher — {today.strftime('%A, %B %d, %Y')}")

    token     = get_token()
    site_id   = get_site_id(token)
    item_id, _ = get_drive_item_id(token, site_id)

    col       = find_today_col(token, site_id, item_id)

    # ── LP Volume (row 64) ────────────────────────────────────────────────────
    vol_addr  = f"{col}{FORECAST_ROW}"
    print(f"📊 Reading LP Volume from {vol_addr}...")
    vol_vals  = read_range(token, site_id, item_id, vol_addr)
    raw_vol   = vol_vals[0][0] if vol_vals and vol_vals[0] else None
    lp_volume = int(float(str(raw_vol).replace(",", ""))) if raw_vol is not None else None
    print(f"✅ LP Volume: {lp_volume:,}" if lp_volume else "⚠️  LP Volume cell empty")

    # ── LP UPH (row 69) ───────────────────────────────────────────────────────
    uph_addr  = f"{col}{UPH_ROW}"
    print(f"📊 Reading LP UPH from {uph_addr}...")
    uph_vals  = read_range(token, site_id, item_id, uph_addr)
    raw_uph   = uph_vals[0][0] if uph_vals and uph_vals[0] else None
    lp_uph    = round(float(str(raw_uph).replace(",", "")), 1) if raw_uph is not None else None
    print(f"✅ LP UPH: {lp_uph}" if lp_uph else "⚠️  LP UPH cell empty")

    write_data_js(lp_volume=lp_volume, lp_uph=lp_uph, cell_vol=vol_addr, cell_uph=uph_addr)


def write_data_js(lp_volume, lp_uph=None, cell_vol: str = "—", cell_uph: str = "—", skipped: bool = False):
    """Write data.js to the dashboard directory."""
    ts = datetime.now().isoformat(timespec="seconds")

    payload = {
        "lpVolume":    lp_volume or 0,
        "lpUPH":       lp_uph or 0,
        "backlog":     0,
        "cellVol":     cell_vol,
        "cellUPH":     cell_uph,
        "lastUpdated": ts,
        "skipped":     skipped,
    }

    js = (
        f"// Auto-generated by fetch_lp_data.py\n"
        f"// Updated: {ts}\n"
        f"window.PHL5_DATA = {json.dumps(payload, indent=2)};\n"
    )

    DATA_JS.write_text(js, encoding="utf-8")
    print(f"💾 Written to {DATA_JS}")


if __name__ == "__main__":
    main()
