"""
create_future_year.py — Stark County Auditor
Create iasWorld FUTURE-YEAR records (next tax year) for every parcel in the
weekly validation portal, by automating the iasWorld "Simple Copy" function.

WHY THIS EXISTS
    iasWorld has no mass "create future-year record" for a list of distinct
    parcels. Simple Copy is one parcel at a time; Advanced Copy copies ONE
    source parcel to MANY new parcels (splits) — neither fits "give each of
    these ~70 distinct parcels its own next-year layer." The system roll would
    roll EVERYTHING, which we don't want yet. So this script does exactly what
    support staff do by hand — open the Multi-Use transaction, load each parcel
    by its PIN, Simple Copy -> next Tax Year -> Create Parcel -> Close — looped
    over the portal's parcel list.

HOW THE iasWorld (EA&T) UI WORKS (confirmed against CAST)
    - Login page (main/Login.aspx) is an Angular shell over the classic form:
      visible fields #wrapperusername / #wrapperpassword (Domain pre-filled
      "stark"), submit via the visible "SIGN IN" button.
    - Parcels can't be deep-linked: Transact.aspx needs a windowId that only
      exists once a transaction is open. So we open the "Multi-Use" transaction
      once (menu -> Transactions -> Multi-Use). That spawns a new window at
      Maintain/Transact.aspx?trans=CAMA_COMBINED which owns the windowId.
    - Inside that window: type the PIN into #txtMaskedPin_seg0 + Enter to load a
      parcel, then click the toolbar "Simple Copy" button. Reuse the same window
      for every parcel.

WHEN TO RUN
    Every Monday, AFTER run_weekly.py produces the week's compare files.
    Prove it in CAST first; only aim at PRODUCTION once proven.

    python create_future_year.py --list-only                 # print worklist, no browser
    python create_future_year.py --env cast --dry-run --diag  # nav + load each parcel, create NOTHING
    python create_future_year.py --env cast --limit 1         # real: one parcel in CAST
    python create_future_year.py --env cast                   # real: full CAST run
    python create_future_year.py --env prod                   # PRODUCTION (must type YES)

OUTPUT
    future_year_log_<date>.csv  (Parcel_ID, status, detail, when)
    status: created | exists | error | dry_run_would_create

NOTE — ROLL THE YEAR FORWARD EACH JANUARY (TARGET_TAX_YEAR below).
"""

import sys
import csv
import argparse
import datetime as _dt
from pathlib import Path

import pandas as pd
import credentials

# Force UTF-8 console output on Windows — cp1252 can't encode the ✓/✗ glyphs
# printed below, and a crash there would abort mid-run (see run_weekly.py).
for _stream in ("stdout", "stderr"):
    _s = getattr(sys, _stream, None)
    if _s is not None and hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

# ── Config ───────────────────────────────────────────────────────────────────
TARGET_TAX_YEAR = 2027                     # bump each January — see module docstring
HERE            = Path(__file__).parent
PORTAL_ROOT     = HERE.parent
MLSCAMA_ROOT    = PORTAL_ROOT / "MLSvsCAMA"
LOGIN_DOMAIN    = "stark"

# Nav label (drawer) — targets the span.left-align-label, which disambiguates the
# menu item from same-named labels elsewhere on the dashboard.
def NAV(label: str) -> str:
    return f"xpath=//span[contains(@class,'left-align-label') and normalize-space(.)='{label}']"

# Confirmed CAST selectors (see docstring). Dialog selectors marked (CONFIRM) are
# best-guess pending a first watched CAST create — --diag dumps the dialog.
SEL = {
    "login_user":   "#wrapperusername",
    "login_pass":   "#wrapperpassword",
    "login_domain": "#wrapperdomain",
    "menu_button":  ["mat-icon:has-text('menu')", "button:has(mat-icon:has-text('menu'))"],
    "drawer_ready": None,   # set at runtime to NAV('Administration')
    "pin_input":    "#txtMaskedPin_seg0",
    # Toolbar Simple Copy: rendered id from <CLIENT_ID>, JS handler from <ONCLICK>CloneParcel()>.
    "simple_copy":  "#btnSimpleCopy",
    "simple_copy_fn": "CloneParcel",
    # Simple Copy dialog renders in an iframe: Maintain/Dialogs/CloneParcel.aspx
    "dialog_frame": "CloneParcel.aspx",
    "copy_year":    "#selYear",     # New Tax Year select IN the dialog frame (defaults to target)
    "copy_create":  "#btnOk",       # <input value="Create Parcel" onclick="cp.OkClick()">
    "copy_close":   ["#btnClose", "input[value='Close']",
                     "xpath=//*[normalize-space(.)='Close']", "xpath=//*[normalize-space(.)='CLOSE']"],
}


# ── Step 1: assemble the parcel worklist (fully testable without iasWorld) ─────

def _norm_pin(v) -> str:
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") else s


def find_latest_week():
    best_folder, best_date = None, None
    for f in MLSCAMA_ROOT.iterdir():
        if not f.is_dir():
            continue
        xlsxs = list(f.glob("value_mismatches_*.xlsx"))
        if not xlsxs:
            continue
        try:
            d = _dt.date.fromisoformat(xlsxs[0].stem.replace("value_mismatches_", ""))
        except ValueError:
            continue
        if best_date is None or d > best_date:
            best_folder, best_date = f, d
    if best_folder is None:
        print("ERROR: no value_mismatches_*.xlsx found under MLSvsCAMA/")
        sys.exit(1)
    return best_folder, best_date.isoformat()


def build_worklist(week_folder: Path, date_str: str) -> list:
    pins, seen = [], set()

    def add(pin):
        pin = _norm_pin(pin)
        if pin and pin.lower() != "nan" and pin not in seen:
            seen.add(pin); pins.append(pin)

    for name in (f"value_mismatches_{date_str}.xlsx", f"perfect_matches_{date_str}.xlsx"):
        path = week_folder / name
        if not path.exists():
            print(f"  (note: {name} not found — skipping)")
            continue
        df = pd.read_excel(path, dtype=str)
        for _, row in df.iterrows():
            add(row.get("Parcel_ID", ""))
            extra = str(row.get("ADDITIONAL_PARCELS", "") or "").strip()
            if extra and extra.lower() != "nan":
                for p in extra.split(","):
                    add(p)
    return pins


# ── Step 2: iasWorld automation (Playwright) ───────────────────────────────────

def _base_url(env: str) -> str:
    key = "IASWORLD_PROD_URL" if env == "prod" else "IASWORLD_CAST_URL"
    url = getattr(credentials, key, None)
    if not url:
        print(f"ERROR: {key} not set in credentials.py")
        sys.exit(1)
    return url


def _dump(pg, tag: str):
    try:
        pg.screenshot(path=str(HERE / f"fy_{tag}.png"), full_page=True)
        (HERE / f"fy_{tag}.html").write_text(pg.content(), encoding="utf-8")
        print(f"    (diag: fy_{tag}.png/.html — {pg.url})")
    except Exception as e:
        print(f"    (diag failed: {e})")


def _click_any(pg, selectors, timeout=4000, force=False):
    from playwright.sync_api import TimeoutError as PWT
    for sel in (selectors if isinstance(selectors, list) else [selectors]):
        try:
            pg.click(sel, timeout=timeout, force=force); return True
        except PWT:
            continue
        except Exception:
            continue
    return False


def login(page, login_url: str, diag: bool) -> bool:
    print("  Logging in to iasWorld...")
    page.goto(login_url, wait_until="domcontentloaded", timeout=45000)
    page.wait_for_timeout(5000)
    if diag:
        _dump(page, "login")
    try:
        page.fill(SEL["login_user"], credentials.IASWORLD_USERNAME, timeout=15000)
        page.fill(SEL["login_pass"], credentials.IASWORLD_PASSWORD, timeout=8000)
        try:
            if not (page.input_value(SEL["login_domain"]) or "").strip():
                page.fill(SEL["login_domain"], LOGIN_DOMAIN)
        except Exception:
            pass
    except Exception:
        _dump(page, "login_fields_missing")
        print("  ✗ Login fields not found."); return False
    # click the VISIBLE 'SIGN IN' button (a hidden ASP #submit shares the text)
    clicked = False
    for btn in page.query_selector_all("button"):
        try:
            if "SIGN IN" in (btn.inner_text() or "").upper() and btn.is_visible():
                btn.click(); clicked = True; break
        except Exception:
            pass
    if not clicked:
        print("  ✗ SIGN IN button not found."); return False
    for _ in range(30):
        page.wait_for_timeout(1000)
        if "login.aspx" not in page.url.lower():
            print("  ✓ Logged in"); return True
    _dump(page, "login_failed")
    print("  ✗ Still on login page — check IASWORLD_USERNAME/PASSWORD."); return False


def open_transaction(page, ctx, diag: bool):
    """Open the Multi-Use transaction; return its (new) page/window. Reused for all PINs."""
    print("  Opening Multi-Use transaction...")
    opened = False
    for _ in range(3):
        _click_any(page, SEL["menu_button"], timeout=2500)
        try:
            page.wait_for_selector(NAV("Administration"), state="visible", timeout=4000)
            opened = True; break
        except Exception:
            page.wait_for_timeout(800)
    if not opened:
        _dump(page, "no_drawer"); print("  ✗ Could not open the navigation drawer."); return None

    def _nav_click(label, timeout=6000):
        loc = page.locator(NAV(label)).first
        try:
            loc.scroll_into_view_if_needed(timeout=4000)
        except Exception:
            pass
        try:
            loc.click(timeout=timeout)
        except Exception:
            loc.click(force=True, timeout=timeout)

    _nav_click("Transactions")                    # expand the Transactions section
    page.wait_for_timeout(1500)
    try:
        page.wait_for_selector(NAV("Multi-Use"), state="visible", timeout=5000)
    except Exception:
        pass
    txn = None
    try:
        with ctx.expect_page(timeout=10000) as pinfo:
            _nav_click("Multi-Use")               # opens the transaction window
        txn = pinfo.value
    except Exception:
        txn = page                                 # some configs open in-place
    txn.wait_for_timeout(6000)
    # confirm we're in the CAMA_COMBINED transaction with the PIN box present
    try:
        txn.wait_for_selector(SEL["pin_input"], timeout=15000)
    except Exception:
        _dump(txn, "no_pin_box")
        print("  ✗ Transaction window has no Parcel ID box (fy_no_pin_box.*)."); return None
    if diag:
        _dump(txn, "transaction")
    print(f"  ✓ Multi-Use open ({txn.url.split('?')[0]})")
    return txn


def load_parcel(txn, pin: str) -> bool:
    """Type the PIN into the ID bar and load it. Returns True if the parcel loaded."""
    try:
        box = txn.query_selector(SEL["pin_input"])
        if not box:
            return False
        box.click()
        # clear then type
        txn.keyboard.press("Control+A")
        txn.keyboard.press("Delete")
        box.type(pin, delay=20)
        txn.keyboard.press("Enter")
        txn.wait_for_timeout(3500)
        return True
    except Exception:
        return False


def future_year_exists(txn, year: int) -> bool:
    """CONFIRM AGAINST CAST — detect an existing {year} record. Conservative False
    until the record-count/tax-year signal is observed (Simple Copy itself reports
    'Record already exists' per-table, so a re-run is not destructive)."""
    return False


def _simple_copy_available(txn) -> bool:
    try:
        if txn.query_selector(SEL["simple_copy"]):
            return True
    except Exception:
        pass
    try:
        return bool(txn.evaluate(f"typeof {SEL['simple_copy_fn']} === 'function'"))
    except Exception:
        return False


def _trigger_simple_copy(txn) -> bool:
    if _click_any(txn, SEL["simple_copy"], timeout=6000):
        return True
    try:
        txn.evaluate(f"{SEL['simple_copy_fn']}()")   # <ONCLICK>CloneParcel()</ONCLICK>
        return True
    except Exception:
        return False


def _dialog_frame(txn):
    """The Simple Copy dialog renders in the CloneParcel.aspx iframe."""
    for fr in txn.frames:
        if SEL["dialog_frame"].lower() in (fr.url or "").lower():
            return fr
    return None


def _cancel_dialog(txn):
    """Cancel/close the CloneParcel dialog if it's still open (cp.cancelClick())."""
    fr = _dialog_frame(txn)
    if not fr:
        return
    for attempt in ("evaluate", "click"):
        try:
            if attempt == "evaluate":
                fr.evaluate("typeof cp!=='undefined' && cp.cancelClick && cp.cancelClick()")
            else:
                fr.click("xpath=//*[normalize-space(.)='Cancel' or normalize-space(.)='CANCEL']", timeout=2000)
            txn.wait_for_timeout(500)
            if not _dialog_frame(txn):
                return
        except Exception:
            pass


def simple_copy(txn, year: int, diag: bool):
    """Simple Copy -> (New Tax Year=year) -> Create Parcel. Returns (status, detail),
    status in {created, exists, error}. iasWorld blocks a duplicate future-year layer,
    so 'no Created-parcel confirmation' means the record already exists."""
    if not _trigger_simple_copy(txn):
        if diag: _dump(txn, "no_simple_copy")
        return "error", "Simple Copy button not found"
    frame = None
    for _ in range(20):
        txn.wait_for_timeout(500)
        frame = _dialog_frame(txn)
        if frame:
            break
    if not frame:
        if diag: _dump(txn, "no_dialog_frame")
        return "error", "Simple Copy dialog did not open"
    if diag: _dump(txn, "copy_dialog")
    # Table checkboxes are pre-set by the site setup form; the year defaults to the
    # target, but set it explicitly for safety. PIN stays as-is (no Generate).
    try:
        frame.select_option(SEL["copy_year"], value=str(year), timeout=5000)
    except Exception:
        pass
    txn.wait_for_timeout(400)
    try:
        frame.click(SEL["copy_create"], timeout=6000)   # #btnOk -> cp.OkClick()
    except Exception:
        if diag: _dump(txn, "no_create_btn")
        _cancel_dialog(txn)
        return "error", "Create Parcel (#btnOk) not clickable"
    # RELIABLE discriminator (confirmed in CAST): after Create Parcel, if the
    # CloneParcel form is still open (#btnOk visible) the create was blocked because
    # the {year} record already exists. If the form closed, the create succeeded and
    # a 'Created parcel … row(s) copied' result renders. (Searching the DOM for
    # 'Created parcel' text is unreliable — a hidden template matches even on exists.)
    def _form_open():
        d = _dialog_frame(txn)
        if not d:
            return False
        try:
            btn = d.query_selector(SEL["copy_create"])
            return bool(btn and btn.is_visible())
        except Exception:
            return False

    created = False
    for _ in range(16):                 # up to ~8s for the form to close on success
        txn.wait_for_timeout(500)
        if not _form_open():
            created = True
            break
    if diag: _dump(txn, "copy_result")

    if not created:                     # duplicate blocked by iasWorld — dismiss the form
        _cancel_dialog(txn)
        return "exists", f"{year} record already exists (create blocked)"

    # Created — read the result summary (a legitimate visible result now), then Close.
    detail = ""
    for fr in [txn.main_frame] + list(txn.frames):
        try:
            for el in fr.query_selector_all("xpath=//*[contains(normalize-space(.),'row(s) copied')]"):
                if el.is_visible():
                    detail = " ".join((el.inner_text() or "").split())[:140]
                    break
        except Exception:
            pass
        if detail:
            break
    for fr in [txn.main_frame] + list(txn.frames):
        for sel in SEL["copy_close"]:
            try:
                if fr.query_selector(sel):
                    fr.click(sel, timeout=2500); txn.wait_for_timeout(500); break
            except Exception:
                pass
    _cancel_dialog(txn)                 # belt-and-suspenders: leave nothing open
    return "created", detail or "created"


def run(pins, env, year, dry_run, diag, headless, limit, log_path):
    from playwright.sync_api import sync_playwright
    login_url = _base_url(env)
    if limit:
        pins = pins[:limit]
    rows, counts = [], {}
    stamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=headless,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
        )
        ctx = browser.new_context(viewport={"width": 1400, "height": 900},
                                  ignore_https_errors=True, accept_downloads=True)
        page = ctx.new_page()
        try:
            if not login(page, login_url, diag):
                sys.exit(1)
            txn = open_transaction(page, ctx, diag)
            if txn is None:
                sys.exit(1)
            # Accept any JS confirm/alert so a create isn't silently cancelled.
            for pg in {page, txn}:
                try:
                    pg.on("dialog", lambda d: d.accept())
                except Exception:
                    pass

            for i, pin in enumerate(pins, 1):
                print(f"  [{i}/{len(pins)}] {pin}")
                status, detail = "error", ""
                try:
                    if not load_parcel(txn, pin):
                        status, detail = "error", "could not load parcel"
                    elif future_year_exists(txn, year):
                        status, detail = "exists", f"{year} record already present"
                    elif dry_run:
                        has_sc = _simple_copy_available(txn)
                        status = "dry_run_would_create" if has_sc else "error"
                        detail = "loaded; Simple Copy available" if has_sc else "Simple Copy button not found"
                        if diag and i == 1:
                            _dump(txn, f"loaded_{pin}")
                    else:
                        status, detail = simple_copy(txn, year, diag and i == 1)
                except Exception as e:
                    status, detail = "error", str(e)[:120]
                counts[status] = counts.get(status, 0) + 1
                rows.append({"Parcel_ID": pin, "status": status, "detail": detail, "when": stamp})
                print(f"      -> {status}" + (f" ({detail})" if detail else ""))
        finally:
            browser.close()

    with open(log_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["Parcel_ID", "status", "detail", "when"])
        w.writeheader(); w.writerows(rows)

    print("\n" + "=" * 60)
    print(f"  FUTURE-YEAR RUN COMPLETE (env={env}, year={year}"
          + (", DRY-RUN" if dry_run else "") + ")")
    for k in ("created", "dry_run_would_create", "exists", "error"):
        if counts.get(k):
            print(f"    {k:24} {counts[k]}")
    print(f"  Log: {log_path}")
    print("=" * 60)


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="Create iasWorld future-year records for portal parcels.")
    ap.add_argument("--env", choices=["cast", "prod"], default="cast")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--diag", action="store_true")
    ap.add_argument("--year", type=int, default=TARGET_TAX_YEAR)
    ap.add_argument("--week", default=None)
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--limit", type=int, default=None, help="only process the first N parcels")
    ap.add_argument("--list-only", action="store_true")
    args = ap.parse_args()

    if args.week:
        week_folder = MLSCAMA_ROOT / args.week
        xlsxs = list(week_folder.glob("value_mismatches_*.xlsx"))
        if not xlsxs:
            print(f"ERROR: no value_mismatches_*.xlsx in {week_folder}"); sys.exit(1)
        date_str = xlsxs[0].stem.replace("value_mismatches_", "")
    else:
        week_folder, date_str = find_latest_week()

    print(f"  Week folder: {week_folder.name}  (date {date_str})")
    pins = build_worklist(week_folder, date_str)
    print(f"  Parcels needing a {args.year} record: {len(pins)}")

    if args.list_only:
        for p in pins:
            print("   ", p)
        return

    if args.env == "prod" and not args.dry_run:
        print(f"\n  *** PRODUCTION: will create {args.year} records for "
              f"{len(pins) if not args.limit else min(args.limit, len(pins))} parcels. ***")
        if input("  Type YES to proceed: ").strip() != "YES":
            print("  Aborted."); return

    log_path = week_folder / f"future_year_log_{date_str}.csv"
    run(pins, args.env, args.year, args.dry_run, args.diag, args.headless, args.limit, log_path)


if __name__ == "__main__":
    main()
