"""
test_mls_login.py — Login-only smoke test for MLS Now (Auth0 SSO).
Mirrors Step 1 of mls_export.py exactly, but stops right after login so we can
verify the updated MLS_PASSWORD works without running the full C3 export.
Exits 0 on success, 1 on failure.
"""

import sys
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# Windows consoles default to cp1252, which can't encode the ✓/✗ glyphs.
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

from credentials import MLS_USERNAME, MLS_PASSWORD

MLS_URL = "https://now.mlsmatrix.com"


def _wait(page, ms=1000):
    page.wait_for_timeout(ms)


def _dump_diag(page, tag):
    out = Path(__file__).parent
    try:
        page.screenshot(path=str(out / f"{tag}.png"), full_page=True)
        (out / f"{tag}.html").write_text(page.content(), encoding="utf-8")
        print(f"    (diagnostic saved: {tag}.png / {tag}.html  — url: {page.url})")
    except Exception as e:
        print(f"    (could not save diagnostic: {e})")


def _try_click(page, selectors, description, timeout=8000):
    for sel in selectors:
        try:
            page.wait_for_selector(sel, timeout=timeout, state="visible")
            page.click(sel)
            print(f"    ✓ {description}")
            return True
        except PWTimeout:
            continue
        except Exception:
            continue
    print(f"    ✗ Could not find: {description}")
    return False


def main():
    print(f"Testing MLS Now login for member {MLS_USERNAME} ...")
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
        )
        context = browser.new_context(viewport={"width": 1280, "height": 900})
        page = context.new_page()

        def _on_login_page():
            c = page.url.lower()
            return ("auth0" in c) or ("login.php" in c) or ("/login" in c)

        logged_in = False
        try:
            for attempt in (1, 2):
                page.goto(MLS_URL, wait_until="domcontentloaded", timeout=30000)

                user_sel = None
                for sel in ["#email", "#loginUsername", "input[type='text']"]:
                    try:
                        page.wait_for_selector(sel, timeout=15000, state="visible")
                        user_sel = sel
                        break
                    except Exception:
                        continue
                if not user_sel:
                    _dump_diag(page, "login_no_field")
                    print("✗ Login form never appeared (no username field).")
                    return 1

                pass_sel = next((s for s in ["#password", "#loginPassword", "input[type='password']"]
                                 if page.query_selector(s)), None)

                page.fill(user_sel, MLS_USERNAME, timeout=5000)
                page.fill(pass_sel,  MLS_PASSWORD, timeout=5000)

                filled = (page.input_value(user_sel) or "").strip()
                if not filled:
                    print(f"    ⚠ username field empty after fill (attempt {attempt}) — retrying")
                    continue

                _try_click(page,
                    ["#btn-login", "#loginButton", "button[type='submit']", "input[type='submit']"],
                    "Login button")

                for _ in range(30):
                    _wait(page, 1000)
                    if not _on_login_page():
                        logged_in = True
                        break
                if logged_in:
                    break
                print(f"    ⚠ still on login portal after submit (attempt {attempt})"
                      + (" — retrying" if attempt == 1 else ""))

            if not logged_in:
                _dump_diag(page, "login_test_failed")
                print("✗ LOGIN FAILED — still on the MLS Now login portal.")
                print("  New password may be wrong, or MLS Now added an extra SSO/MFA step.")
                return 1

            print(f"✓ LOGIN OK — reached: {page.url}")
            return 0
        finally:
            browser.close()


if __name__ == "__main__":
    sys.exit(main())
