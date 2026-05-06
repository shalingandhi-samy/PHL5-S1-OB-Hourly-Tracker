#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = ["playwright", "python-dotenv"]
# ///
"""
PHL5 Cookie Extractor — Gscope only.
Opens system Edge so Walmart SSO works. Writes GSCOPE_TOKEN + GSCOPE_COOKIE into .env.
Re-run any morning the Gscope token expires (~8–24 hrs).

NOTE: DRAX auth is handled natively by the qa-kitten scheduler task (browser SSO).
      No DRAX cookie needed here.
"""
import asyncio
import re
import sys
from pathlib import Path

HERE     = Path(__file__).parent
ENV_FILE = HERE / ".env"
ENV_EX   = HERE / ".env.example"

GSCOPE_URL = "https://gscope.walmart.com/mfe/distributorCapacity"


# ── helpers ───────────────────────────────────────────────────────────────────

def _load_env() -> str:
    if ENV_FILE.exists():
        return ENV_FILE.read_text(encoding="utf-8")
    if ENV_EX.exists():
        print("  No .env found — seeding from .env.example")
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


# ── Gscope ────────────────────────────────────────────────────────────────────

async def extract_gscope(context) -> tuple[str | None, str | None]:
    """Returns (gateway_token, full_cookie_string) or (None, None)."""
    print("\n── Gscope ───────────────────────────────────────────────────────────")
    page = await context.new_page()
    try:
        await page.goto(GSCOPE_URL, wait_until="domcontentloaded", timeout=60_000)

        # If SSO didn't auto-auth, give the user 90 s to log in manually
        if "login" in page.url.lower() or "coreid" in page.url.lower():
            print("  ⚠️  SSO redirect — please log in to Gscope in the browser window.")
            print("  ⏳ Waiting up to 90 seconds...")
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
    print("\n🐱 PHL5 Cookie Extractor — Gscope")
    print("  Opening Edge — SSO will fire automatically on Walmart network.")
    print("  (DRAX handled by qa-kitten — no cookie needed here)\n")

    from playwright.async_api import async_playwright

    async with async_playwright() as pw:
        try:
            browser = await pw.chromium.launch(
                headless=False, channel="msedge", args=["--start-maximized"]
            )
        except Exception:
            browser = await pw.chromium.launch(headless=False, args=["--start-maximized"])
        context = await browser.new_context(no_viewport=True)

        gscope_token, gscope_cs = await extract_gscope(context)

        await browser.close()

    print("\n── Patching .env ────────────────────────────────────────────────────")
    env = _load_env()

    if gscope_token:
        env = _upsert(env, "GSCOPE_TOKEN",  gscope_token)
        env = _upsert(env, "GSCOPE_COOKIE", gscope_cs or "")
        print(f"  GSCOPE_TOKEN  → {len(gscope_token)} chars")
        print(f"  GSCOPE_COOKIE → {len(gscope_cs or '')} chars")
        _save_env(env)
        print("\n✅ Done! Re-run tomorrow before shift.")
    else:
        print("  ⚠️  GSCOPE_TOKEN not captured — check Gscope login and re-run.")
    print()


if __name__ == "__main__":
    asyncio.run(main())
