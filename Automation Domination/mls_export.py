"""
mls_export.py — Export MLS data from MLS Matrix using Playwright
Logs into now.mlsmatrix.com, runs the CAMAvsMLS saved search, exports C3 CSV,
then auto-converts the CSV to Excel (eliminates the manual conversion step).
"""

import sys
import time
import pandas as pd
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

from credentials import MLS_USERNAME, MLS_PASSWORD

MLS_URL = "https://now.mlsmatrix.com"


def _wait(page, ms=2000):
    page.wait_for_timeout(ms)


def _dump_diag(page, tag):
    """Write a screenshot + page HTML next to this script so a selector that
    broke against MLS Now's portal can be fixed without a blind re-run."""
    out = Path(__file__).parent
    try:
        page.screenshot(path=str(out / f"{tag}.png"), full_page=True)
        (out / f"{tag}.html").write_text(page.content(), encoding="utf-8")
        print(f"    (diagnostic saved: {tag}.png / {tag}.html  — url: {page.url})")
    except Exception as e:
        print(f"    (could not save diagnostic: {e})")


def _try_click(page, selectors, description, timeout=8000):
    """Try a list of selectors until one works. Returns True on success."""
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


def export_mls(csv_path: Path, xlsx_path: Path) -> pd.DataFrame:
    """
    Full MLS Matrix export workflow via Playwright.
    Saves the raw CSV to csv_path and the converted Excel to xlsx_path.
    Returns a DataFrame of the MLS data.
    """
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
        )
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1280, "height": 900},
        )
        page = context.new_page()

        try:
            # ── Step 1: Login ──────────────────────────────────────────────────
            print("  Step 1: Logging in to MLS Matrix...")
            # MLS Now moved to an Auth0/SAML SSO portal (login.php). Field IDs:
            #   Member ID = #email, password = #password, submit = #btn-login.
            # The form is JS-driven, so we must WAIT for #email to exist and
            # confirm the fields actually filled before submitting — otherwise a
            # race submits an empty form and bounces straight back to login
            # (looks identical to a bad password). Retry the whole sequence once.
            def _on_login_page():
                c = page.url.lower()
                return ("auth0" in c) or ("login.php" in c) or ("/login" in c)

            logged_in = False
            for attempt in (1, 2):
                page.goto(MLS_URL, wait_until="domcontentloaded", timeout=30000)

                # Wait for the real username field to be interactive (not a fixed sleep).
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
                    print("  ✗ Login form never appeared (no username field).")
                    sys.exit(1)

                pass_sel = next((s for s in ["#password", "#loginPassword", "input[type='password']"]
                                 if page.query_selector(s)), None)

                page.fill(user_sel, MLS_USERNAME, timeout=5000)
                page.fill(pass_sel,  MLS_PASSWORD, timeout=5000)

                # Confirm the fields are actually populated before submitting.
                filled = (page.input_value(user_sel) or "").strip()
                if not filled:
                    print(f"    ⚠ username field empty after fill (attempt {attempt}) — retrying")
                    continue

                _try_click(page,
                    ["#btn-login", "#loginButton", "button[type='submit']", "input[type='submit']"],
                    "Login button")

                # The SSO chain (login.php → auth0 → SAML → main.pl) can take
                # well over 6s. Poll for the redirect to leave the login portal
                # rather than checking once after a fixed sleep (that race caused
                # false "still on login" failures even when login succeeded).
                for _ in range(30):           # up to ~30s
                    _wait(page, 1000)
                    if not _on_login_page():
                        logged_in = True
                        break
                if logged_in:
                    break
                print(f"    ⚠ still on login portal after submit (attempt {attempt})"
                      + (" — retrying" if attempt == 1 else ""))

            if not logged_in:
                _dump_diag(page, "login_failed")
                print("  ✗ Login did not complete — still on the MLS Now login portal.")
                print("    Check MLS_USERNAME/MLS_PASSWORD in credentials.py, or whether")
                print("    MLS Now now requires an extra SSO/MFA step.")
                print(f"    Diagnostic written next to mls_export.py (login_failed.*).")
                sys.exit(1)
            print("  ✓ Logged in")

            # ── Step 2: Click Matrix icon ──────────────────────────────────────
            print("  Step 2: Clicking Matrix icon...")
            clicked = _try_click(page,
                ["#1", "img[src*='CoreLogicMatrix']", "img[title='Matrix']", "img[alt='Matrix']"],
                "Matrix icon")
            if not clicked:
                # Last resort: find any image with 'matrix' in attributes
                images = page.query_selector_all("img")
                for img in images:
                    src   = img.get_attribute("src")   or ""
                    title = img.get_attribute("title") or ""
                    alt   = img.get_attribute("alt")   or ""
                    iid   = img.get_attribute("id")    or ""
                    if any("matrix" in x.lower() for x in [src, title, alt]) or iid == "1":
                        img.click()
                        clicked = True
                        print("    ✓ Matrix icon (search)")
                        break
                if not clicked:
                    _dump_diag(page, "matrix_icon_missing")
                    print("  ✗ Matrix icon not found — cannot continue")
                    print("    (You ARE logged in, but the Matrix launch element changed.")
                    print("     Inspect matrix_icon_missing.html to pin the new selector.)")
                    sys.exit(1)
            _wait(page, 5000)

            # ── Step 3: Navigate to Saved Searches ────────────────────────────
            print("  Step 3: Navigating to Saved Searches...")

            # First try: hover My Matrix dropdown then click link
            clicked = False
            my_matrix = None
            for xpath in [
                "//span[@class='text-uppercase' and contains(text(),'My Matrix')]",
                "//span[contains(text(),'My Matrix')]",
                "//*[contains(text(),'My Matrix')]",
            ]:
                try:
                    my_matrix = page.query_selector(f"xpath={xpath}")
                    if my_matrix:
                        break
                except Exception:
                    continue

            if my_matrix:
                my_matrix.hover()
                _wait(page, 2000)
                clicked = _try_click(page,
                    [
                        "a[href='/Matrix/SavedSearches']",
                        "a[href='/Matrix/SavedSearches.aspx']",
                        "a[href*='SavedSearches']",
                        "xpath=//a[contains(text(),'My Saved Searches')]",
                        "xpath=//a[contains(text(),'Saved Searches')]",
                        "xpath=//span[contains(text(),'My Saved Searches')]/parent::a",
                    ],
                    "My Saved Searches",
                    timeout=5000)

            # Fallback: navigate directly to the Saved Searches URL
            if not clicked:
                print("    Trying direct navigation to Saved Searches...")
                current_url = page.url
                base = current_url.split("/Matrix/")[0] if "/Matrix/" in current_url else MLS_URL
                for path in ["/Matrix/SavedSearches", "/Matrix/SavedSearches.aspx"]:
                    try:
                        page.goto(base + path, wait_until="domcontentloaded", timeout=15000)
                        _wait(page, 2000)
                        print("    ✓ Navigated directly to Saved Searches")
                        clicked = True
                        break
                    except Exception:
                        continue

            if not clicked:
                print("  ✗ Could not reach Saved Searches page")
                sys.exit(1)
            _wait(page, 3000)

            # ── Step 4: Click CAMAvsMLS saved search ───────────────────────────
            print("  Step 4: Opening CAMAvsMLS search...")
            clicked = _try_click(page,
                [
                    "#m_sscv_m_lvSS_ssdi_2538498_ssv_m_aNameToggleAppearance",
                    "xpath=//a[contains(text(),'CAMAvsMLS')]",
                    "xpath=//a[contains(@id,'aNameToggleAppearance')]",
                ],
                "CAMAvsMLS search link")
            if not clicked:
                print("  ✗ CAMAvsMLS search not found")
                sys.exit(1)
            _wait(page, 2000)

            # ── Step 5: Click Results ──────────────────────────────────────────
            print("  Step 5: Loading results...")
            clicked = _try_click(page,
                [
                    "#m_sscv_m_lvSS_ssdi_2538498_ssv_d_ssdv_m_btnFullSearch",
                    "xpath=//a[text()='Results']",
                    "xpath=//input[@value='Results']",
                ],
                "Results button")
            if not clicked:
                print("  ✗ Results button not found")
                sys.exit(1)
            _wait(page, 6000)

            # ── Step 6: Check All ──────────────────────────────────────────────
            print("  Step 6: Selecting all properties...")
            clicked = _try_click(page,
                [
                    "#m_lnkCheckAllLink",
                    "a[title='Check All']",
                    "xpath=//a[contains(text(),'Check All')]",
                ],
                "Check All")
            if not clicked:
                print("  ✗ Check All not found")
                sys.exit(1)
            _wait(page, 2000)

            # ── Step 7: Click Export ───────────────────────────────────────────
            print("  Step 7: Opening export dialog...")
            clicked = _try_click(page,
                [
                    "span.linkIcon.icon_export",
                    "xpath=//span[contains(@class,'icon_export')]/parent::a",
                    "xpath=//a[contains(text(),'Export')]",
                ],
                "Export button")
            if not clicked:
                print("  ✗ Export button not found")
                sys.exit(1)
            _wait(page, 3000)

            # ── Step 8: Select C3 format ───────────────────────────────────────
            print("  Step 8: Selecting C3 format...")
            try:
                page.select_option("#m_ddExport", value="ud11476", timeout=8000)
                print("    ✓ C3 format selected (by value)")
            except PWTimeout:
                try:
                    page.select_option("#m_ddExport", label="C3", timeout=5000)
                    print("    ✓ C3 format selected (by label)")
                except PWTimeout:
                    print("  ✗ Could not select C3 format")
                    sys.exit(1)
            _wait(page, 1000)

            # ── Step 9: Click final Export and capture download ────────────────
            print("  Step 9: Downloading CSV...")
            try:
                with page.expect_download(timeout=60000) as dl_info:
                    _try_click(page,
                        ["#m_btnExport", "a[id='m_btnExport']"],
                        "final Export button")
                download = dl_info.value
                download.save_as(str(csv_path))
                print(f"  ✓ Downloaded: {csv_path.name}")
            except PWTimeout:
                print("  ✗ Download timed out")
                sys.exit(1)

        finally:
            browser.close()

    # ── Convert CSV → Excel ────────────────────────────────────────────────────
    print("  Converting CSV to Excel...")
    df = pd.read_csv(csv_path, dtype=str)  # read as strings first
    df = df.fillna("0").replace("", "0")

    # Coerce numeric-looking columns to numbers.
    # Skip columns where any value has a leading zero — those are ID/code
    # fields (e.g. Parcel Number "00106948") that must stay as strings.
    for col in df.columns:
        original = df[col].astype(str)
        has_leading_zero = original.str.match(r'^0\d').any()
        if has_leading_zero:
            continue
        converted = pd.to_numeric(df[col], errors="coerce")
        if converted.notna().sum() / max(len(df), 1) > 0.8:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df.to_excel(xlsx_path, index=False)
    print(f"  ✓ Converted to Excel: {xlsx_path.name}")
    return df
