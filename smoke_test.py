#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = ["msal", "requests", "python-dotenv", "urllib3"]
# ///
"""
PHL5 Hourly Tracker — Smoke Test
Checks every auth source and API endpoint without blocking for login.
Run this to diagnose what's working before / after a cookie refresh.
"""
import json
import os
import sys
import urllib3
from datetime import date, datetime
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent / ".env")
except ImportError:
    pass

import msal
import requests

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HERE     = Path(__file__).parent
EG_CACHE = HERE / ".token_cache_endgame.json"
GR_CACHE = HERE / ".token_cache_graph.json"

TENANT_ID = "3cbcc3d3-094d-4006-9849-0d11d61f484d"
EG_CLIENT = "777d5c3a-eb0c-43d5-b30a-798a7eb9d15e"
GR_CLIENT = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
FC_ID     = "3124"

DRAX_BASE   = os.getenv("DRAX_BASE_URL",   "https://drax.walmart.com")
GSCOPE_BASE = os.getenv("GSCOPE_BASE_URL", "https://gscope.walmart.com")
DRAX_COOKIE   = os.getenv("DRAX_COOKIE",   "")
GSCOPE_COOKIE = os.getenv("GSCOPE_COOKIE", "")
GSCOPE_TOKEN  = os.getenv("GSCOPE_TOKEN",  "")

PASS  = "\033[92m PASS \033[0m"
FAIL  = "\033[91m FAIL \033[0m"
WARN  = "\033[93m WARN \033[0m"
INFO  = "\033[94m INFO \033[0m"

results: list[tuple[str, str, bool | None, str]] = []  # (system, check, ok, detail)


def row(system: str, check: str, ok: bool | None, detail: str = "") -> None:
    tag = PASS if ok is True else (FAIL if ok is False else WARN)
    results.append((system, check, ok, detail))
    print(f"  {'[OK]' if ok is True else '[!!]' if ok is False else '[--]':<6} {system:<18} {check:<30} {detail}")


# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "="*70)
print("  PHL5 Hourly Tracker — Smoke Test")
print(f"  {datetime.now().strftime('%A %b %d %Y  %H:%M:%S')}")
print("="*70 + "\n")

# ── 1. .env / config ─────────────────────────────────────────────────────────
print("── Config ──────────────────────────────────────────────────────────")
env_file = HERE / ".env"
row("Config", ".env file exists",      env_file.exists(),
    str(env_file) if env_file.exists() else "Create from .env.example!")
row("Config", "DRAX_COOKIE set",       bool(DRAX_COOKIE),
    f"{len(DRAX_COOKIE)} chars" if DRAX_COOKIE else "MISSING — paste from browser DevTools")
row("Config", "GSCOPE auth set",       bool(GSCOPE_COOKIE or GSCOPE_TOKEN),
    "GSCOPE_TOKEN set" if GSCOPE_TOKEN else
    ("GSCOPE_COOKIE set" if GSCOPE_COOKIE else "MISSING — paste from gscope.walmart.com"))

# ── 2. MSAL token caches ──────────────────────────────────────────────────────
print("\n── MSAL Token Caches ───────────────────────────────────────────────")

def _msal_cached_token(client_id: str, scopes: list[str], cache_file: Path,
                        label: str) -> tuple[bool | None, str]:
    """Try silent token acquire. Returns (ok, detail). Never blocks for login."""
    if not cache_file.exists():
        return None, "No cache file — run the fetcher once to authenticate"
    try:
        cache = msal.SerializableTokenCache()
        cache.deserialize(cache_file.read_text())
        app = msal.PublicClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            token_cache=cache,
        )
        accounts = app.get_accounts()
        if not accounts:
            return None, "Cache exists but no accounts — token may have expired"
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            # Decode exp claim without extra deps
            parts = result["access_token"].split(".")
            if len(parts) >= 2:
                import base64
                pad = 4 - len(parts[1]) % 4
                payload = json.loads(base64.urlsafe_b64decode(parts[1] + "=" * pad))
                exp = datetime.fromtimestamp(payload.get("exp", 0))
                return True, f"Valid — expires {exp.strftime('%b %d %H:%M')}"
            return True, "Valid (exp unknown)"
        if result and "error" in result:
            return False, f"Expired/revoked — {result.get('error_description','')[:60]}"
        return None, "Silent acquire returned no result — may need re-auth"
    except Exception as e:
        return False, f"Error: {e}"

eg_ok, eg_detail = _msal_cached_token(
    EG_CLIENT, [f"{EG_CLIENT}/.default"], EG_CACHE, "Endgame")
gr_ok, gr_detail = _msal_cached_token(
    GR_CLIENT,
    ["https://graph.microsoft.com/Files.Read.All",
     "https://graph.microsoft.com/Sites.Read.All"],
    GR_CACHE, "MS Graph")

row("Endgame MSAL",  "Cached token",  eg_ok, eg_detail)
row("MS Graph MSAL", "Cached token",  gr_ok, gr_detail)

# ── 3. Endgame API ───────────────────────────────────────────────────────────
print("\n── Endgame API ─────────────────────────────────────────────────────")

def _eg_headers() -> dict:
    h = {"Accept": "application/json"}
    if eg_ok and eg_detail.startswith("Valid"):
        # re-acquire silently — we know it works
        cache = msal.SerializableTokenCache()
        cache.deserialize(EG_CACHE.read_text())
        app = msal.PublicClientApplication(
            EG_CLIENT,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            token_cache=cache,
        )
        result = app.acquire_token_silent(
            [f"{EG_CLIENT}/.default"], account=app.get_accounts()[0])
        if result and "access_token" in result:
            h["Authorization"] = f"Bearer {result['access_token']}"
    return h

if eg_ok:
    eg_headers = _eg_headers()
    # pt-status
    try:
        r = requests.get(
            f"https://status-api.endgame-status.prod.k8s.walmart.net/v3/{FC_ID}",
            headers=eg_headers, timeout=10, verify=False)
        row("Endgame", "pt-status /v3", r.status_code == 200,
            f"HTTP {r.status_code}" + (f" — backlog keys: {list(r.json().keys())[:5]}"
                                        if r.status_code == 200 else ""))
    except Exception as e:
        row("Endgame", "pt-status /v3", False, str(e)[:60])

    # /cutoffs
    try:
        r = requests.get(
            f"https://status-api.endgame-status.prod.k8s.walmart.net/{FC_ID}/cutoffs",
            headers=eg_headers,
            params={"date": date.today().isoformat()},
            timeout=10, verify=False)
        has_groups = "cutOffGroups" in (r.json() if r.status_code == 200 else {})
        row("Endgame", "/cutoffs (OTS)", r.status_code == 200 and has_groups,
            f"HTTP {r.status_code}" + (f" — cutOffGroups present: {has_groups}"
                                        if r.status_code == 200 else ""))
    except Exception as e:
        row("Endgame", "/cutoffs (OTS)", False, str(e)[:60])
else:
    row("Endgame", "pt-status /v3",    None, "Skipped — no valid MSAL token")
    row("Endgame", "/cutoffs (OTS)",   None, "Skipped — no valid MSAL token")

# ── 4. DRAX ──────────────────────────────────────────────────────────────────
print("\n── DRAX ────────────────────────────────────────────────────────────")
if DRAX_COOKIE:
    try:
        date_str = date.today().strftime("%Y-%m-%d")
        r = requests.get(
            f"{DRAX_BASE}/building_overview/",
            params={
                "date_hour_after":  f"{date_str}+07:00",
                "date_hour_before": f"{date_str}+23:59",
                "area": "Outbound",
            },
            headers={
                "Cookie":   DRAX_COOKIE,
                "Referer":  f"{DRAX_BASE}/",
                "Accept":   "text/html,*/*",
                "User-Agent": "Mozilla/5.0",
            },
            timeout=15, verify=False,
        )
        authed      = r.status_code not in (401, 403)
        has_data    = "Stationary Picking" in r.text if authed else False
        row("DRAX", "building_overview",  authed and has_data,
            f"HTTP {r.status_code}" + (" — Stationary Picking row found" if has_data
                                        else " — auth OK but no data yet" if authed
                                        else " — COOKIE EXPIRED"))
    except Exception as e:
        row("DRAX", "building_overview", False, str(e)[:60])
else:
    row("DRAX", "building_overview", False,
        "DRAX_COOKIE not set — paste from browser DevTools → Network → Cookie header")

# ── 5. Gscope ────────────────────────────────────────────────────────────────
print("\n── Gscope ──────────────────────────────────────────────────────────")
if GSCOPE_COOKIE or GSCOPE_TOKEN:
    bearer = GSCOPE_TOKEN.replace("Bearer ", "").strip() if GSCOPE_TOKEN else None
    if not bearer and GSCOPE_COOKIE:
        for part in GSCOPE_COOKIE.split(";"):
            p = part.strip()
            if p.startswith("gateway_token="):
                bearer = p[len("gateway_token="):]
    gs_headers = {
        "Accept": "application/json",
        "Referer": f"{GSCOPE_BASE}/mfe/distributorCapacity",
        "Origin": "https://gscope.walmart.com",
        "X-Requested-With": "XMLHttpRequest",
    }
    if GSCOPE_COOKIE:
        gs_headers["Cookie"] = GSCOPE_COOKIE
    if bearer:
        gs_headers["Authorization"] = f"Bearer {bearer}"

    try:
        r = requests.get(
            "https://gscope.walmart.com/api/gateway/fcap/getFCDetail",
            params={"distributorId": "31341", "poolId": "1"},
            headers=gs_headers, timeout=10, verify=False)
        authed = r.status_code not in (401, 403)
        has_data = bool(r.json().get("payload") or r.json().get("plannedCapacity")) \
                   if r.status_code == 200 else False
        row("Gscope", "FCAP getFCDetail (1P)", authed and r.status_code == 200,
            f"HTTP {r.status_code}" + (" — capacity data found" if has_data
                                        else " — auth OK, unexpected shape" if authed
                                        else " — TOKEN/COOKIE EXPIRED"))
    except Exception as e:
        row("Gscope", "FCAP getFCDetail (1P)", False, str(e)[:60])
else:
    row("Gscope", "FCAP getFCDetail (1P)", False,
        "GSCOPE_COOKIE / GSCOPE_TOKEN not set")

# ── 6. data.js ───────────────────────────────────────────────────────────────
print("\n── data.js ─────────────────────────────────────────────────────────")
data_js = HERE / "data.js"
if data_js.exists():
    try:
        txt   = data_js.read_text(encoding="utf-8")
        start = txt.find("{"); end = txt.rfind("}") + 1
        d     = json.loads(txt[start:end])
        ts    = d.get("lastUpdated", "unknown")
        age_s = (datetime.now() - datetime.fromisoformat(ts)).total_seconds() \
                if ts != "unknown" else 9999
        age_h = age_s / 3600
        row("data.js", "Parseable",    True,  f"lastUpdated={ts}")
        row("data.js", "Freshness",    age_h < 2,
            f"{age_h:.1f} h old" + (" — STALE" if age_h >= 2 else " — fresh"))
        row("data.js", "otsData[20]",  len(d.get("otsData", [])) == 20,
            f"len={len(d.get('otsData', []))}")
        row("data.js", "capacity[3]",  len(d.get("capacity", [])) == 3,
            f"len={len(d.get('capacity', []))}")
    except Exception as e:
        row("data.js", "Parseable", False, str(e)[:60])
else:
    row("data.js", "Exists", False, "File missing — run fetch_lp_data.py first")

# ── 7. Task Scheduler ─────────────────────────────────────────────────────────
print("\n── Task Scheduler ──────────────────────────────────────────────────")
import subprocess
for task in ("PHL5-LP-Fetch", "PHL5-Hourly-Fetch"):
    try:
        out = subprocess.check_output(
            ["schtasks", "/query", "/tn", task, "/fo", "LIST"],
            text=True, stderr=subprocess.DEVNULL)
        status = "Ready" if "Ready" in out else ("Running" if "Running" in out else "Unknown")
        nxt    = next((l.split(":", 1)[1].strip()
                       for l in out.splitlines() if "Next Run" in l), "?")
        row("Scheduler", task, status in ("Ready", "Running"),
            f"{status} — next: {nxt}")
    except Exception:
        row("Scheduler", task, False, "Task not found in scheduler!")

# ── Summary ───────────────────────────────────────────────────────────────────
print("\n" + "="*70)
fails  = [r for r in results if r[2] is False]
warns  = [r for r in results if r[2] is None]
passes = [r for r in results if r[2] is True]
print(f"  PASSED: {len(passes)}   WARNINGS: {len(warns)}   FAILED: {len(fails)}")

if fails:
    print("\n  ACTION REQUIRED:")
    for _, check, _, detail in fails:
        print(f"    \u2022 {check}: {detail}")
if warns:
    print("\n  NEEDS ATTENTION:")
    for _, check, _, detail in warns:
        print(f"    \u2022 {check}: {detail}")

print()

# ── Expiry Advisory ───────────────────────────────────────────────────────────
print("── What Can Expire (Read This!) ────────────────────────────────────")
advisories = [
    ("DRAX_COOKIE",        "EXPIRES",   "~8–24 hrs",  "Refresh from drax.walmart.com → F12 → Network → Cookie header → paste in .env"),
    ("GSCOPE_TOKEN/COOKIE","EXPIRES",   "~8–24 hrs",  "Refresh from gscope.walmart.com → F12 → Application → Cookies → gateway_token"),
    ("Endgame MSAL",       "SAFE",      "90 days",    "Auto-refreshes silently. If it expires, run fetch_hourly.py manually once to re-auth via device code."),
    ("MS Graph MSAL",      "SAFE",      "90 days",    "Same as above — run fetch_lp_data.py manually once to re-auth."),
    ("Task Scheduler",     "SAFE",      "Permanent",  "Runs as long as you are LOGGED IN. Laptop sleep is fine; full logoff is not."),
    ("Battery mode",       "WATCH OUT", "—",          "Tasks are set to NOT run on battery. Plug in during shift hours!"),
]
fmt = "  {:<22} {:<10} {:<12} {}"
print(fmt.format("COMPONENT", "STATUS", "LIFESPAN", "WHAT TO DO"))
print("  " + "-"*66)
for comp, status, life, action in advisories:
    icon = "OK " if status == "SAFE" else ("!!!" if status == "EXPIRES" else " ! ")
    print(fmt.format(comp, f"[{icon}] {status}", life, action))

print("\n  Bottom line:")
print("  Cookies (DRAX + Gscope) need refreshing DAILY — before each shift.")
print("  Everything else is fully automated once authenticated.\n")
