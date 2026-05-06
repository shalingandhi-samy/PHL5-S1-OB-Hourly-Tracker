#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "msal",
#   "requests",
#   "websocket-client",
# ]
# ///
"""
PHL5 Auto-Screenshot -> Teams
Runs at :20 of each hour, Mon-Thu, 8:20 AM - 5:20 PM via Task Scheduler.

How it works:
  1. Launch Edge headless with --remote-debugging-port (CDP)
  2. Load index.html, wait for Alpine.js to settle
  3. Full-page JPEG screenshot via CDP (no msedgedriver needed)
  4. Post image to Teams group chat via MS Graph hostedContents API
"""

import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import base64
import json
import subprocess
import time
import os
import msal
import requests
import websocket
from datetime import datetime
from pathlib import Path

# ── Paths ─────────────────────────────────────────────────────────────────────
HERE        = Path(__file__).parent
INDEX_HTML  = HERE / "index.html"
TOKEN_CACHE = HERE / ".token_cache_teams.json"
LOG_FILE    = HERE / "screenshot_teams.log"
EDGE_EXE    = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
CDP_PORT    = 9223   # avoid clashing with any existing Edge debug session

# ── MS Graph ──────────────────────────────────────────────────────────────────
TENANT_ID   = "3cbcc3d3-094d-4006-9849-0d11d61f484d"
GRAPH_CLIENT = "14d82eec-204b-4c2f-b7e8-296a70dab67e"   # MS Graph Explorer
GRAPH_SCOPES = [
    "https://graph.microsoft.com/Chat.ReadWrite",
]

# ── Target chat ───────────────────────────────────────────────────────────────
CHAT_ID     = "19:c4ce152ed2fa4944851d081c3ab01c67@thread.v2"  # I appreciate your feedback!

# ── Logging ───────────────────────────────────────────────────────────────────
def log(msg: str) -> None:
    ts  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


# ── MS Graph auth (MSAL device code) ─────────────────────────────────────────
def get_token() -> str:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE.exists():
        cache.deserialize(TOKEN_CACHE.read_text(encoding="utf-8"))

    app = msal.PublicClientApplication(
        GRAPH_CLIENT,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )

    accounts = app.get_accounts()
    result   = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0]) if accounts else None

    if not result:
        flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
        print("\n" + flow["message"] + "\n")
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description')}")

    TOKEN_CACHE.write_text(cache.serialize(), encoding="utf-8")
    return result["access_token"]


# ── Edge CDP screenshot ───────────────────────────────────────────────────────
def take_screenshot() -> bytes:
    """Launch Edge headless, load index.html, return full-page JPEG bytes."""
    url = INDEX_HTML.as_uri()
    log(f"Launching Edge headless -> {url}")

    proc = subprocess.Popen([
        EDGE_EXE,
        f"--remote-debugging-port={CDP_PORT}",
        "--headless",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-extensions",
        "--hide-scrollbars",
        f"--window-size=1920,1080",
        url,
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    try:
        # Wait for Edge to start and page to load + Alpine.js to init
        time.sleep(4)

        # Get the websocket debug URL
        r = requests.get(f"http://localhost:{CDP_PORT}/json", timeout=5)
        targets = r.json()
        ws_url  = next(
            t["webSocketDebuggerUrl"] for t in targets
            if t.get("type") == "page"
        )

        ws = websocket.create_connection(ws_url, timeout=10)
        _id = 0

        def cdp(method, params=None):
            nonlocal _id
            _id += 1
            ws.send(json.dumps({"id": _id, "method": method, "params": params or {}}))
            return json.loads(ws.recv())

        # Wait a bit more for Alpine.js reactive rendering
        cdp("Runtime.evaluate", {"expression": "new Promise(r => setTimeout(r, 2500))", "awaitPromise": True})

        # Get full page dimensions
        dims_raw = cdp("Runtime.evaluate", {
            "expression": "JSON.stringify({w: document.body.scrollWidth, h: document.body.scrollHeight})"
        })
        dims = json.loads(dims_raw["result"]["result"]["value"])
        pw, ph = dims["w"], dims["h"]
        log(f"Page size: {pw} x {ph}")

        # Set device metrics to full page size at 2x DPR
        cdp("Emulation.setDeviceMetricsOverride", {
            "width":             pw,
            "height":            ph,
            "deviceScaleFactor": 2,
            "mobile":            False,
        })

        # Capture full-page JPEG
        shot = cdp("Page.captureScreenshot", {
            "format":               "jpeg",
            "quality":              88,
            "captureBeyondViewport": True,
            "clip": {"x": 0, "y": 0, "width": pw, "height": ph, "scale": 2},
        })
        ws.close()

        img_bytes = base64.b64decode(shot["result"]["data"])
        log(f"Screenshot captured: {len(img_bytes):,} bytes")
        return img_bytes

    finally:
        proc.terminate()
        proc.wait()


# ── Post to Teams ─────────────────────────────────────────────────────────────
def post_to_teams(img_bytes: bytes, token: str) -> None:
    now_label = datetime.now().strftime("%I:%M %p  %a %b %d")
    img_b64   = base64.b64encode(img_bytes).decode()

    payload = {
        "body": {
            "contentType": "html",
            "content": (
                f"<p><b>PHL5 S1 OB Hourly Tracker &mdash; {now_label}</b></p>"
                f'<img src="../hostedContents/1/$value">'
            ),
        },
        "hostedContents": [
            {
                "@microsoft.graph.temporaryId": "1",
                "contentBytes":                  img_b64,
                "contentType":                   "image/jpeg",
            }
        ],
    }

    resp = requests.post(
        f"https://graph.microsoft.com/v1.0/chats/{CHAT_ID}/messages",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        },
        json=payload,
        timeout=30,
        verify=False,
    )
    resp.raise_for_status()
    log(f"Posted to Teams OK  (HTTP {resp.status_code})")


# ── Main ──────────────────────────────────────────────────────────────────────
def main() -> None:
    log("--- PHL5 Auto-Screenshot start ---")
    try:
        token     = get_token()
        img_bytes = take_screenshot()
        post_to_teams(img_bytes, token)
        log("--- Done ---")
    except Exception as exc:
        log(f"ERROR: {exc}")
        raise


if __name__ == "__main__":
    main()
