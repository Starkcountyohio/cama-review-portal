"""
run_datapull.py — Monday pipeline, DATA-PULL HALF ONLY (steps 1–3).
Runs CAMA extract → MLS export → compare, writing the 4 Excel files to the
week folder, then STOPS. Zillow photos / photo review / build / push are the
operator's half (run_weekly.py handles the full chain). Mirrors run_weekly.py
steps 1–3 exactly so the outputs are identical to a normal Monday run.
"""

import sys
from datetime import date
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HERE        = Path(__file__).parent
PORTAL_ROOT = HERE.parent
MLSCAMA_ROOT = PORTAL_ROOT / "MLSvsCAMA"

sys.path.insert(0, str(HERE))
sys.path.insert(0, str(PORTAL_ROOT))

from cama_extract import extract_cama, get_week_dates
from mls_export   import export_mls
from compare      import run_comparison


def folder_name_from_date(d: date) -> str:
    return f"{d.month}-{d.day:02d}-{str(d.year)[2:]}"


def date_str_from_date(d: date) -> str:
    return d.strftime("%Y-%m-%d")


def run(run_date: date = None):
    run_date = run_date or date.today()
    start_date, end_date = get_week_dates(run_date)
    folder_name = folder_name_from_date(run_date)
    date_str    = date_str_from_date(run_date)

    print("=" * 70)
    print("  AUTOMATION DOMINATION — Data Pull (steps 1–3 only)")
    print("=" * 70)
    print(f"  Run date:        {run_date}  ({run_date.strftime('%A')})")
    print(f"  Sale date range: {start_date} to {end_date} (exclusive)")
    print(f"  Output folder:   MLSvsCAMA/{folder_name}/")
    print("=" * 70)

    week_folder = MLSCAMA_ROOT / folder_name
    week_folder.mkdir(parents=True, exist_ok=True)

    # ── Step 1: CAMA extract ────────────────────────────────────────────────
    print(f"\n[1/3] CAMA Extract")
    cama_csv = week_folder / f"cama_{date_str}.csv"
    df_cama  = extract_cama(cama_csv, start_date=start_date, end_date=end_date)

    # ── Step 2: MLS export ──────────────────────────────────────────────────
    print(f"\n[2/3] MLS Export")
    mls_csv  = week_folder / f"mls_{date_str}.csv"
    mls_xlsx = week_folder / f"mls_{date_str}.xlsx"
    df_mls   = export_mls(mls_csv, mls_xlsx)

    # ── Step 3: Compare ─────────────────────────────────────────────────────
    print(f"\n[3/3] CAMA vs MLS Comparison")
    counts = run_comparison(df_mls, df_cama, week_folder, date_str)
    print(f"\n  Summary:")
    print(f"    Value mismatches:  {counts['mismatches']}")
    print(f"    Perfect matches:   {counts['perfects']}")
    print(f"    Missing in CAMA:   {counts['missing_cama']}")
    print(f"    Missing in MLS:    {counts['missing_mls']}")
    likely = counts.get('likely_matches', 0)
    if likely:
        print(f"    Likely parcel typos: {likely}  ← review "
              f"likely_parcel_mismatches_{date_str}.xlsx")

    print("\n" + "=" * 70)
    print("  DATA PULL COMPLETE — ready for Zillow")
    print("=" * 70)
    print(f"\n  Week folder: MLSvsCAMA/{folder_name}/")
    print(f"  Next (operator): run Zillow photos → photo review → build → push")
    print(f"  (run_weekly.py picks up the full chain; or run ZillowPhotos manually)")
    return counts


if __name__ == "__main__":
    run()
