"""
run_weekly.py — Automation Domination: Monday morning pipeline
Stark County Auditor — MLS vs CAMA Review Portal

Steps (run in order):
  1. CAMA extract     → Oracle → cama_YYYY-MM-DD.csv
  2. MLS export       → MLS Matrix → mls_YYYY-MM-DD.csv + .xlsx
  3. Compare          → 4 Excel files (mismatches, perfects, missing_cama, missing_mls)
  4. Zillow photos    → Photos_New/ + Photos_New_Portal/  (requires manual CAPTCHA)
  5. Photo verify     → Retry-loop on parcels still missing a -1.jpg; user can
                        retry, open URLs in browser to confirm, or skip remaining
     (photo-review gate — always pauses for the agent-face/wrong-house scan)
  6. Photo-upload CSV → Photo Upload M-D-YYYY.csv (iasWorld Document Loader),
                        Taxyr = create_future_year.TARGET_TAX_YEAR, incl. manual photos
  7. Build portal     → review_portal.html + archive/YYYY-WNN.json
  8. Push to GitHub

Manual steps (not automated — done in iasWorld after this script + create_future_year.py):
  - Sale Tab Mass Update
  - MassEntrance
  - iasWorld photo upload (use the Photo Upload CSV from step 6, onto the future-year records)
"""

import sys
import subprocess
from datetime import date, timedelta
from pathlib import Path

# Force UTF-8 output on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ── Paths ──────────────────────────────────────────────────────────────────────
HERE         = Path(__file__).parent                            # Automation Domination/
PORTAL_ROOT  = HERE.parent                                      # Portal Builder/
MLSCAMA_ROOT = PORTAL_ROOT / "MLSvsCAMA"                        # MLSvsCAMA/
ZILLOW_SCRIPT = PORTAL_ROOT / "ZillowPhotos" / "download_zillow_photos.py"
BUILD_SCRIPT  = PORTAL_ROOT / "build_portal.py"

# ── Imports from this project ──────────────────────────────────────────────────
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(PORTAL_ROOT))

from cama_extract import extract_cama, get_week_dates
from mls_export   import export_mls
from compare      import run_comparison
import credentials

# build_portal functions imported directly (skip its interactive main())
import build_portal as bp
import make_photo_upload as mpu
import photo_checks
import datetime as _dt


# ── Date helpers ───────────────────────────────────────────────────────────────

def folder_name_from_date(d: date) -> str:
    """Format date as M-DD-YY for folder naming (matches existing convention)."""
    return f"{d.month}-{d.day:02d}-{str(d.year)[2:]}"


def date_str_from_date(d: date) -> str:
    """Format date as YYYY-MM-DD for file naming."""
    return d.strftime("%Y-%m-%d")


# ── Photo verification helper ──────────────────────────────────────────────────

def _load_photo_targets(mismatches_xlsx, perfects_xlsx):
    """Return list of {parcel_id, address, city, state, zip} dicts from both Excels."""
    import pandas as pd
    rows = []
    for xlsx in (mismatches_xlsx, perfects_xlsx):
        if not xlsx.exists():
            continue
        df = pd.read_excel(xlsx).drop_duplicates('Parcel_ID', keep='first')
        for _, r in df.iterrows():
            rows.append({
                'parcel_id': str(r['Parcel_ID']),
                'address':   str(r['Address']).strip(),
                'city':      str(r['City']).strip(),
                'state':     'OH' if pd.isna(r.get('State')) else str(r['State']).strip(),
                'zip':       str(int(r['Zip'])) if not pd.isna(r['Zip']) else '',
            })
    return rows


def verify_photos(week_folder, mismatches_xlsx, perfects_xlsx):
    """
    Loop until every parcel has Photos_New/<id>-1.jpg, or the user attests that
    the remaining failures truly have no photos available.

    Interactive prompt each pass:
      [R] retry — re-run the downloader (auto-skips already-saved parcels)
      [O] open the Zillow search URL for each missing parcel in the browser
      [S] skip remaining (user verified) and continue to portal build
      [Q] quit pipeline before push
    """
    import webbrowser
    import urllib.parse

    photos_dir = week_folder / "Photos_New"
    targets = _load_photo_targets(mismatches_xlsx, perfects_xlsx)
    if not targets:
        print("  (no parcels to verify)")
        return

    pass_num = 1
    while True:
        missing = [t for t in targets
                   if not (photos_dir / f"{t['parcel_id']}-1.jpg").exists()]

        if not missing:
            print(f"\n  ✓ All {len(targets)} parcels have an exterior photo.")
            return

        bar = "─" * 60
        print(f"\n  ┌{bar}")
        print(f"  │ Photo verification — pass {pass_num}")
        print(f"  │ {len(missing)} of {len(targets)} parcels still missing a photo:")
        for i, t in enumerate(missing, 1):
            print(f"  │   [{i}] {t['parcel_id']:>10}  "
                  f"{t['address']}, {t['city']}, {t['state']} {t['zip']}")
        print(f"  └{bar}")

        print("\n  Options:")
        print("    [R] Retry — re-run downloader on the missing parcels")
        print("    [O] Open Zillow URLs in browser (visual review)")
        print("    [S] Skip remaining and continue to portal build (no photos available)")
        print("    [Q] Quit pipeline (abort before push)")
        try:
            choice = input("  Choice: ").strip().lower()
        except EOFError:
            print("\n  (non-interactive — skipping verification)")
            return

        if choice == 'r':
            print()
            for xlsx_file, label in [
                (mismatches_xlsx, "value_mismatches"),
                (perfects_xlsx,   "perfect_matches"),
            ]:
                if not xlsx_file.exists():
                    continue
                print(f"  Re-running downloader for {label}...")
                subprocess.run(
                    [sys.executable, str(ZILLOW_SCRIPT),
                     str(xlsx_file), str(photos_dir)],
                    cwd=str(PORTAL_ROOT),
                )
            pass_num += 1
        elif choice == 'o':
            for t in missing:
                full = f"{t['address']}, {t['city']}, {t['state']} {t['zip']}"
                url = f"https://www.zillow.com/homes/{urllib.parse.quote(full)}_rb/"
                webbrowser.open(url)
            print(f"\n  Opened {len(missing)} URL(s) in the default browser.")
        elif choice == 's':
            print(f"\n  Skipping remaining {len(missing)} parcel(s) "
                  f"(user verified no photos available). Continuing to portal build.")
            return
        elif choice == 'q':
            sys.exit("\n  Pipeline aborted by user before portal build/push.")
        else:
            print("  Invalid choice — please pick R, O, S, or Q.")


# ── Manual photo-review checkpoint (before build) ───────────────────────────────

def review_photos_checkpoint(week_folder):
    """Unconditional hard stop before the portal is built.

    verify_photos() only pauses when a photo is MISSING and only checks that a
    -1.jpg EXISTS — it can't catch a realtor headshot, the wrong house, or a bad
    crop. This gate always pauses so the exteriors get eyeballed before anything
    is built or pushed.

    A swap now only needs to be saved ONCE, into either folder: the sync step
    immediately after this gate propagates it to the other one and prints what it
    moved. Saving into both by hand is still harmless (identical bytes = no-op).
    """
    import webbrowser

    photos_dir  = week_folder / "Photos_New"
    portal_dir  = week_folder / "Photos_New_Portal"
    mapping_csvs = sorted(photos_dir.glob("*.csv")) if photos_dir.is_dir() else []

    bar = "─" * 60
    print(f"\n  ┌{bar}")
    print(f"  │ PHOTO REVIEW — required before build")
    print(f"  │ Review exteriors for agent faces / wrong house / bad crops.")
    print(f"  │ Save a swap as <parcel>-1.jpg in EITHER folder below — the next")
    print(f"  │ step syncs it to the other one and reports what it moved.")
    print(f"  │")
    print(f"  │   Photos_New:        {photos_dir}")
    print(f"  │   Photos_New_Portal: {portal_dir}")
    if mapping_csvs:
        print(f"  │   Mapping CSV(s):")
        for c in mapping_csvs:
            print(f"  │     - {c.name}")
    print(f"  └{bar}")

    while True:
        print("\n  Options:")
        print("    [O] Open both photo folders in Explorer")
        print("    [build] Photos reviewed — build the portal and push")
        print("    [Q] Quit pipeline (abort before build/push)")
        try:
            choice = input("  Choice: ").strip().lower()
        except EOFError:
            # No stdin: do NOT silently build/push unreviewed photos.
            sys.exit("\n  Non-interactive session — aborting before build "
                     "(photo review gate could not be confirmed).")

        if choice == 'o':
            for d in (photos_dir, portal_dir):
                if d.is_dir():
                    webbrowser.open(d.as_uri())
            print("  Opened photo folder(s) in Explorer.")
        elif choice == 'build':
            print("  Photos confirmed — continuing to build.")
            return
        elif choice == 'q':
            sys.exit("\n  Pipeline aborted by user at photo-review gate "
                     "(before build/push).")
        else:
            print("  Type 'build' to proceed, 'o' to open folders, or 'q' to quit.")


# ── Main pipeline ──────────────────────────────────────────────────────────────

def run(run_date: date = None, start_date=None, end_date=None):
    run_date = run_date or date.today()
    if start_date is None or end_date is None:
        start_date, end_date = get_week_dates(run_date)
    folder_name = folder_name_from_date(run_date)
    date_str    = date_str_from_date(run_date)

    print("=" * 70)
    print("  AUTOMATION DOMINATION — Weekly Pipeline")
    print("=" * 70)
    print(f"  Run date:        {run_date}  ({run_date.strftime('%A')})")
    print(f"  Sale date range: {start_date} to {end_date} (exclusive)")
    print(f"  Output folder:   MLSvsCAMA/{folder_name}/")
    print("=" * 70)

    # Create output folder
    week_folder = MLSCAMA_ROOT / folder_name
    week_folder.mkdir(parents=True, exist_ok=True)

    # ── Step 1: CAMA extract ───────────────────────────────────────────────────
    print(f"\n[1/8] CAMA Extract")
    cama_csv = week_folder / f"cama_{date_str}.csv"
    df_cama  = extract_cama(cama_csv, start_date=start_date, end_date=end_date)

    # ── Step 2: MLS export ────────────────────────────────────────────────────
    print(f"\n[2/8] MLS Export")
    mls_csv  = week_folder / f"mls_{date_str}.csv"
    mls_xlsx = week_folder / f"mls_{date_str}.xlsx"
    df_mls   = export_mls(mls_csv, mls_xlsx)

    # ── Step 3: Compare ────────────────────────────────────────────────────────
    print(f"\n[3/8] CAMA vs MLS Comparison")
    counts = run_comparison(df_mls, df_cama, week_folder, date_str)
    print(f"\n  Summary:")
    print(f"    Value mismatches:  {counts['mismatches']}")
    print(f"    Perfect matches:   {counts['perfects']}")
    print(f"    Missing in CAMA:   {counts['missing_cama']}")
    print(f"    Missing in MLS:    {counts['missing_mls']}")
    likely = counts.get('likely_matches', 0)
    if likely:
        print(f"    Likely parcel typos: {likely}  ← review "
              f"likely_parcel_mismatches_{date_str}.xlsx before portal build")

    # ── Step 4: Zillow photos ──────────────────────────────────────────────────
    print(f"\n[4/8] Zillow Photos  (requires manual CAPTCHA solving)")
    photos_dir = week_folder / "Photos_New"

    mismatches_xlsx = week_folder / f"value_mismatches_{date_str}.xlsx"
    perfects_xlsx   = week_folder / f"perfect_matches_{date_str}.xlsx"

    # Run value_mismatches first (more photos needed), then perfect_matches
    for xlsx_file, label in [
        (mismatches_xlsx, "value_mismatches"),
        (perfects_xlsx,   "perfect_matches"),
    ]:
        if not xlsx_file.exists():
            print(f"  Skipping {label} — file not found")
            continue
        print(f"\n  Downloading photos for {label}...")
        result = subprocess.run(
            [sys.executable, str(ZILLOW_SCRIPT),
             str(xlsx_file), str(photos_dir)],
            cwd=str(PORTAL_ROOT),
        )
        if result.returncode != 0:
            print(f"  ⚠  Zillow downloader exited with code {result.returncode}")

    # ── Step 5: Photo verification ─────────────────────────────────────────────
    print(f"\n[5/8] Photo Verification")
    verify_photos(week_folder, mismatches_xlsx, perfects_xlsx)

    # ── Photo-review gate (always pauses, even when all photos present) ─────────
    review_photos_checkpoint(week_folder)

    # ── Primary-photo sync, immediately after the review gate ──────────────────
    # Swaps made during the review are propagated between Photos_New and
    # Photos_New_Portal here, so the operator no longer has to write each fix into
    # two directories from memory (the actual cause of six weeks of one-folder
    # swaps — see photo_checks.sync_primaries()).
    print(f"\n[Photo sync] Primary photos across both folders")
    photo_checks.report_sync(photo_checks.sync_primaries(week_folder))

    # ── Automated photo checks, AFTER the human review gate ────────────────────
    # review_photos_checkpoint() asks the operator to eyeball the exteriors, but
    # can't catch a swap that only landed in Photos_New — filenames and counts
    # look identical in both folders, so only a content hash finds it. Run here so
    # manual swaps made during the review are validated before anything is built.
    print(f"\n[Photo checks] Pre-build verification")
    if not photo_checks.gate(week_folder, mismatches_xlsx, perfects_xlsx):
        sys.exit(1)

    # ── Step 6: Photo-upload CSV (iasWorld Document Loader) ─────────────────────
    # Emitted AFTER the review gate so manual photo swaps are captured. Taxyr is
    # set from create_future_year.TARGET_TAX_YEAR so the photos attach to the
    # future-year record. Used later in the MANUAL iasWorld photo upload (which is
    # done after create_future_year.py has created those records). Non-fatal: a
    # failure here must not block the portal build/push.
    print(f"\n[6/8] Photo-Upload CSV")
    try:
        mpu.generate(week_folder, date_str)
    except Exception as e:
        print(f"  ⚠  Photo-upload CSV not generated ({e}). Build continues — create it "
              f"manually with make_photo_upload.py before the iasWorld photo upload.")

    # ── Step 7: Build portal ───────────────────────────────────────────────────
    print(f"\n[7/8] Building Portal")

    portal_photos_dir = week_folder / "Photos_New_Portal"
    output_html       = PORTAL_ROOT / "review_portal.html"

    # Read Excel files
    mismatches_rows = bp.xlsx_to_json(mismatches_xlsx) if mismatches_xlsx.exists() else []
    perfects_rows   = bp.xlsx_to_json(perfects_xlsx)   if perfects_xlsx.exists()   else []

    # Load photos
    photo_map = bp.load_photos(portal_photos_dir) if portal_photos_dir.is_dir() else {}

    # Week label
    monday    = run_date - _dt.timedelta(days=run_date.weekday())
    week_label = f"Week of {monday.strftime('%B %d, %Y')}"
    week_key   = bp.get_iso_week_key(run_date)

    print(f"  Week label: {week_label}")

    # Build HTML
    generated_at = _dt.datetime.now().strftime('%B %d, %Y at %I:%M %p')
    html = bp.build_html(
        mismatches_rows, perfects_rows, photo_map, week_label, generated_at,
        api_key="",  # Key stored in staff browsers via localStorage — never embedded in HTML
        github_pages_base=bp.GITHUB_PAGES_BASE,
        shared_api_url=bp.SHARED_API_URL,
    )

    output_html.write_text(html, encoding="utf-8")
    size_mb = output_html.stat().st_size / (1024 * 1024)
    print(f"  ✓ review_portal.html  ({size_mb:.1f} MB)")

    # Write version.json (build marker for the staff reload-prompt). MUST carry the
    # SAME generated_at baked into the HTML above, or staff get a permanent false
    # "newer version available" banner that no reload can clear.
    import json as _json
    version_path = PORTAL_ROOT / "version.json"
    version_path.write_text(_json.dumps({
        "generatedAt": generated_at,
        "weekLabel":   week_label,
        "weekKey":     week_key,
    }), encoding="utf-8")
    print(f"  ✓ version.json  (build marker: {generated_at})")

    # Save archive
    archive_path = bp.save_archive(
        mismatches_rows, perfects_rows,
        week_key, week_label, PORTAL_ROOT,
    )
    print(f"  ✓ {archive_path.relative_to(PORTAL_ROOT)}")

    # ── Step 7: Push to GitHub ────────────────────────────────────────────────
    print(f"\n[8/8] Pushing to GitHub")
    import subprocess as _sp
    try:
        _sp.run(["git", "add", "review_portal.html", "version.json", str(archive_path.relative_to(PORTAL_ROOT))],
                cwd=str(PORTAL_ROOT), check=True)
        _sp.run(["git", "commit", "-m", f"Portal update — {week_label}"],
                cwd=str(PORTAL_ROOT), check=True)
        _sp.run(["git", "push"], cwd=str(PORTAL_ROOT), check=True)
        print(f"  ✓ Pushed review_portal.html + archive/{week_key}.json to GitHub")
    except _sp.CalledProcessError as e:
        print(f"  ⚠  GitHub push failed (you can push manually): {e}")

    # ── Done ───────────────────────────────────────────────────────────────────
    print("\n" + "=" * 70)
    print("  PIPELINE COMPLETE")
    print("=" * 70)
    print(f"\n  Weekly files:  MLSvsCAMA/{folder_name}/")
    print(f"  Portal:        review_portal.html")
    print(f"  Archive:       archive/{week_key}.json")
    print()
    print("  Manual steps remaining in iasWorld:")
    print("    - Sale Tab Mass Update")
    print("    - MassEntrance")
    print(f"    - Photo upload via Document Loader CSV in Photos_New/")
    print()
    try:
        input("Press Enter to close...")
    except EOFError:
        pass


if __name__ == "__main__":
    run()
