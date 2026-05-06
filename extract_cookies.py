#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = ["playwright", "python-dotenv"]
# ///
"""
PHL5 Cookie Extractor — DRAX + Gscope
Opens a real browser (non-headless) so Walmart SSO works.
Writes DRAX_COOKIE, GSCOPE_COOKIE, GSCOPE_TOKEN into .env.
Re-run any morning cookies expire (~8–24 hrs).
"""
import asyncio
import re
import sys
from pathlib import Path

HERE     = Path(__file__).parent
ENV_FILE = HERE / ".env"
ENV_EX   = HERE / ".env.example"

DRAX_URL   = (
    "https://drax.walmart.com/building_overview/"
    "?date_hour_after=2026-05-06+07:00"
    "&date_hour_before=2026-05-06+23:59"
    "&area=Outbound"
)
GSCOPE_URL = "https://gscope.walmart.com/mfe/distributorCapacity"


# ── helpers ───────────────────────────────────────────────────────────────────

def _load_env() -> str:
    if ENV_FILE.exists():
        return ENV_FILE.read_text(encoding="utf-8")
    if ENV_EX.exists():
        print(f"  No .env found — seeding from .env.example")
        return ENV_EX.read_text(encoding="utf-8")
    return ""


def _upsert(env: str, key: str, value: str) -> str:
    """Insert or replace a KEY=value line in an env file string."""
    line = f"{key}={value}"
    if re.search(rf"^{key}=", env, re.MULTILINE):
        return re.sub(rf"^{key}=.*$", line, env, flags=re.MULTILINE)
    return env.rstrip() + f"\n{line}\n"


def _save_env(content: str) -> None:
    ENV_FILE.write_text(content, encoding="utf-8")
    print(f"  ✅ .env written → {ENV_FILE}")


def _cookie_str(cookies: list[dict], domain_hint: str) -> str:
    return "; ".join(
        f"{c['name']}={c['value']}"
        for c in cookies
        if domain_hint in c.get("domain", "")
    )


# ── DRAX ──────────────────────────────────────────────────────────────────────

async def extract_drax(context) -> str | None:
    print("\n── DRAX ─────────────────────────────────────────────────────────────")
    page = await context.new_page()
    try:
        resp = await page.goto(DRAX_URL, wait_until="networkidle", timeout=60_000)
        status = resp.status if resp else "?"
        title  = await page.title()
        print(f"  HTTP {status}  |  {title}")

        if "login" in page.url.lower():
            print("  ❌ Redirected to login — check VPN / Eagle WiFi")
            return None

        cookies = await context.cookies(["https://drax.walmart.com"])
        cs = _cookie_str(cookies, "drax.walmart.com")
        has_session = any(c["name"] == "sessionid" for c in cookies)

        body = await page.inner_text("body")
        has_data = "Stationary Picking" in body

        print(f"  🍪 {len(cookies)} cookies  |  sessionid={'✅' if has_session else '❌'}  "
              f"|  Stationary Picking={'✅' if has_data else '❌ (no data yet — check shift time)'}")
        return cs if has_session else None

    except Exception as e:
        print(f"  ❌ Error: {e}")
        return None
    finally:
        await page.close()


# ── Gscope ────────────────────────────────────────────────────────────────────

async def extract_gscope(context) -> tuple[str | None, str | None]:
    """Returns (gateway_token, full_cookie_string) or (None, None)."""
    print("\n── Gscope ───────────────────────────────────────────────────────────")
    page = await context.new_page()
    try:
        await page.goto(GSCOPE_URL, wait_until="domcontentloaded", timeout=60_000)

        # If SSO didn't auto-auth, give the user 90 s to log in manually
        if "login" in page.url.lower() or "coreid" in page.url.lower():
            print("  ⚠️  SSO redirect detected — please log in to Gscope in the browser window.")
            print("  ⏳ Waiting up to 90 seconds for you to complete login...")
            try:
                await page.wait_for_url(
                    lambda url: "gscope.walmart.com" in url and "login" not in url,
                    timeout=90_000,
                )
            except Exception:
                print("  ❌ Timed out waiting for Gscope login")
                return None, None

        await page.wait_for_load_state("networkidle", timeout=30_000)

        cookies = await context.cookies(["https://gscope.walmart.com"])
        token   = next((c["value"] for c in cookies if c["name"] == "gateway_token"), None)
        cs      = _cookie_str(cookies, "gscope.walmart.com")
        title   = await page.title()

        print(f"  🍪 {len(cookies)} cookies  |  gateway_token={'✅' if token else '❌'}  |  {title}")
        return token, (cs if token else None)

    except Exception as e:
        print(f"  ❌ Error: {e}")
        return None, None
    finally:
        await page.close()


# ── main ──────────────────────────────────────────────────────────────────────

async def main() -> None:
    print("\n🐱 PHL5 Cookie Extractor")
    print("  Opening browser — SSO will fire automatically on Walmart network.\n")

    from playwright.async_api import async_playwright

    async with async_playwright() as pw:
        # Use system Edge (always on Windows) — avoids downloading Chromium binaries.
        # Falls back to bundled Chromium if Edge isn't found (shouldn't happen on Win).
        try:
            browser = await pw.chromium.launch(
                headless=False, channel="msedge", args=["--start-maximized"]
            )
        except Exception:
            browser = await pw.chromium.launch(headless=False, args=["--start-maximized"])
        context = await browser.new_context(no_viewport=True)

        drax_cs              = await extract_drax(context)
        gscope_token, gscope_cs = await extract_gscope(context)

        await browser.close()

    # ── patch .env ────────────────────────────────────────────────────────────
    print("\n── Patching .env ────────────────────────────────────────────────────")
    env = _load_env()
    ok  = True

    if drax_cs:
        env = _upsert(env, "DRAX_COOKIE", drax_cs)
        print(f"  DRAX_COOKIE   → {len(drax_cs)} chars")
    else:
        print("  ⚠️  DRAX_COOKIE not updated — check DRAX connection")
        ok = False

    if gscope_token:
        env = _upsert(env, "GSCOPE_TOKEN",  gscope_token)
        env = _upsert(env, "GSCOPE_COOKIE", gscope_cs or "")
        print(f"  GSCOPE_TOKEN  → {len(gscope_token)} chars")
        print(f"  GSCOPE_COOKIE → {len(gscope_cs or '')} chars")
    else:
        print("  ⚠️  GSCOPE_TOKEN not updated — check Gscope login")
        ok = False

    _save_env(env)

    print("\n" + ("✅ All cookies captured! Re-run this each morning before shift."
                  if ok else
                  "⚠️  Some cookies missing — check warnings above and re-run."))
    print()


if __name__ == "__main__":
    asyncio.run(main())
