"""
run_build.py — Monday pipeline, BUILD HALF ONLY (steps 6–7, NO push).
Regenerates the iasWorld photo-upload CSV, then builds review_portal.html +
version.json + the weekly archive from the already-downloaded photos. Mirrors
run_weekly.py steps 6–7 exactly. Does NOT git commit/push — that's a separate,
deliberate step. Run for a week whose photos are already in Photos_New_Portal.
"""

import sys
import json as _json
import datetime as _dt
from datetime import date
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HERE         = Path(__file__).parent
PORTAL_ROOT  = HERE.parent
MLSCAMA_ROOT = PORTAL_ROOT / "MLSvsCAMA"

sys.path.insert(0, str(HERE))
sys.path.insert(0, str(PORTAL_ROOT))

import build_portal as bp
import make_photo_upload as mpu
import photo_checks


def folder_name_from_date(d: date) -> str:
    return f"{d.month}-{d.day:02d}-{str(d.year)[2:]}"


def run(run_date: date = None, skip_photo_checks: bool = False,
        strict_photos: bool = False):
    run_date = run_date or date.today()
    folder_name = folder_name_from_date(run_date)
    date_str    = run_date.strftime("%Y-%m-%d")
    week_folder = MLSCAMA_ROOT / folder_name

    print("=" * 70)
    print("  AUTOMATION DOMINATION — Build (steps 6–7, no push)")
    print("=" * 70)
    print(f"  Run date:      {run_date}  ({run_date.strftime('%A')})")
    print(f"  Week folder:   MLSvsCAMA/{folder_name}/")
    print("=" * 70)

    mismatches_xlsx = week_folder / f"value_mismatches_{date_str}.xlsx"
    perfects_xlsx   = week_folder / f"perfect_matches_{date_str}.xlsx"

    # ── Primary-photo sync: runs BEFORE the gate ─────────────────────────────
    # A manual swap only ever landed in one folder because nothing propagated it
    # — see sync_primaries() for why the long-assumed "silent copy failure" was a
    # misdiagnosis. Doing this ahead of the gate means the gate now verifies the
    # sync rather than being the only thing standing between a missed swap and
    # the live portal.
    print(f"\n[Photo sync] Primary photos across both folders")
    photo_checks.report_sync(photo_checks.sync_primaries(week_folder))

    # ── Photo gate: runs BEFORE the CSV and the build ────────────────────────
    # Deliberately ahead of step 6 — a bad photo must not reach either the
    # iasWorld upload CSV or the portal. See photo_checks.py for what's a hard
    # failure vs. a warning, and why.
    print(f"\n[Photo checks] Pre-build verification")
    if not photo_checks.gate(week_folder, mismatches_xlsx, perfects_xlsx,
                             skip=skip_photo_checks, strict=strict_photos):
        sys.exit(1)

    # ── Step 6: Photo-upload CSV (regenerate to pick up manual photo fixes) ──
    print(f"\n[6/7] Photo-Upload CSV (regenerate)")
    try:
        out = mpu.generate(week_folder, date_str)
        print(f"  ✓ Regenerated: {out.name}")
    except Exception as e:
        print(f"  ⚠  Photo-upload CSV not generated ({e}). Build continues.")

    # ── Step 7: Build portal ────────────────────────────────────────────────
    print(f"\n[7/7] Building Portal")
    portal_photos_dir = week_folder / "Photos_New_Portal"
    output_html       = PORTAL_ROOT / "review_portal.html"

    mismatches_rows = bp.xlsx_to_json(mismatches_xlsx) if mismatches_xlsx.exists() else []
    perfects_rows   = bp.xlsx_to_json(perfects_xlsx)   if perfects_xlsx.exists()   else []
    photo_map = bp.load_photos(portal_photos_dir) if portal_photos_dir.is_dir() else {}

    monday     = run_date - _dt.timedelta(days=run_date.weekday())
    week_label = f"Week of {monday.strftime('%B %d, %Y')}"
    week_key   = bp.get_iso_week_key(run_date)
    print(f"  Week label: {week_label}")
    print(f"  Photos loaded: {len(photo_map)}")

    generated_at = _dt.datetime.now().strftime('%B %d, %Y at %I:%M %p')
    html = bp.build_html(
        mismatches_rows, perfects_rows, photo_map, week_label, generated_at,
        api_key="",
        github_pages_base=bp.GITHUB_PAGES_BASE,
        shared_api_url=bp.SHARED_API_URL,
    )
    output_html.write_text(html, encoding="utf-8")
    size_mb = output_html.stat().st_size / (1024 * 1024)
    print(f"  ✓ review_portal.html  ({size_mb:.1f} MB)")

    version_path = PORTAL_ROOT / "version.json"
    version_path.write_text(_json.dumps({
        "generatedAt": generated_at,
        "weekLabel":   week_label,
        "weekKey":     week_key,
    }), encoding="utf-8")
    print(f"  ✓ version.json  (build marker: {generated_at})")

    archive_path = bp.save_archive(
        mismatches_rows, perfects_rows, week_key, week_label, PORTAL_ROOT,
    )
    print(f"  ✓ {archive_path.relative_to(PORTAL_ROOT)}")

    print("\n" + "=" * 70)
    print("  BUILD COMPLETE — NOT pushed")
    print("=" * 70)
    print(f"\n  To publish, commit + push:")
    print(f"    review_portal.html")
    print(f"    version.json")
    print(f"    archive/{week_key}.json")
    return {"week_key": week_key, "week_label": week_label,
            "size_mb": size_mb, "photos": len(photo_map),
            "mismatches": len(mismatches_rows), "perfects": len(perfects_rows),
            "archive": str(archive_path.relative_to(PORTAL_ROOT))}


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--skip-photo-checks", action="store_true",
                    help="build even if the photo gate finds hard failures "
                         "(escape hatch — the failures it catches have all been real)")
    ap.add_argument("--strict-photos", action="store_true",
                    help="also treat photo WARNINGS (e.g. parcels with no primary "
                         "photo) as blocking")
    args = ap.parse_args()
    print(run(skip_photo_checks=args.skip_photo_checks,
              strict_photos=args.strict_photos))
