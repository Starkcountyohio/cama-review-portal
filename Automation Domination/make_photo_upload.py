"""
make_photo_upload.py — Stark County Auditor
Generate the iasWorld "Document Loader" photo-upload CSV for a week, by scanning
the week's Photos_New folder. Replaces the old hand-built spreadsheet.

WHY THIS EXISTS
    The weekly photo-upload CSV used to come from a template (Photo Upload
    Templet.xlsx) and was built before any manual photo downloads, so (a) manual
    photos got missed, and (b) every row defaulted to Taxyr=2026. Starting with
    the future-year workflow (2026-07), photos must attach to the FUTURE tax-year
    record (2027) created by create_future_year.py — not the current year.

    This script fixes both permanently: it scans Photos_New (so every photo
    present — auto-downloaded or manual — is included) and stamps Taxyr with the
    same TARGET_TAX_YEAR that create_future_year.py uses, so the two stay in sync.
    Roll the year forward once, in create_future_year.py, and this follows.

WHAT IT EMITS
    One row per primary exterior photo (<parcel>-1.jpg) — matching how prior
    weeks' uploads worked (only the primary is pushed to iasWorld). Columns:
      Filename,FileSize,Rank,Parid,Jur,Taxyr,Photo Category,Card,Title,Notes,
      Photo Capture Date,SubjectXCoord,GPS_LAT,SubjectYCoord,GPS_LONG

USAGE
    python make_photo_upload.py                 # latest week -> Photo Upload M-D-YYYY.csv
    python make_photo_upload.py --week 7-13-26  # a specific week folder
    python make_photo_upload.py --year 2027     # override the tax year
    python make_photo_upload.py --out some.csv  # override the output path
    python make_photo_upload.py --dry-run       # print summary, write nothing
"""

import sys
import csv
import re
import argparse
import datetime as _dt
from pathlib import Path

# Force UTF-8 console output on Windows — cp1252 can't encode the ✓/✗ glyphs below.
for _stream in ("stdout", "stderr"):
    _s = getattr(sys, _stream, None)
    if _s is not None and hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

# Single source of truth for the tax year + week discovery: reuse create_future_year.
from create_future_year import TARGET_TAX_YEAR, MLSCAMA_ROOT, find_latest_week

HEADER = ["Filename", "FileSize", "Rank", "Parid", "Jur", "Taxyr",
          "Photo Category", "Card", "Title", "Notes", "Photo Capture Date",
          "SubjectXCoord", "GPS_LAT", "SubjectYCoord", "GPS_LONG"]

JUR  = "000"
CARD = "1"
PRIMARY_RE = re.compile(r"(\d+)-1\.jpg$", re.IGNORECASE)


def _date_from_folder(week_folder: Path, date_str: str) -> _dt.date:
    """Prefer the ISO date_str from find_latest_week; fall back to the folder name."""
    try:
        return _dt.date.fromisoformat(date_str)
    except (ValueError, TypeError):
        m, d, y = week_folder.name.split("-")           # e.g. 7-13-26
        return _dt.date(2000 + int(y), int(m), int(d))


def build_rows(photos_dir: Path, year: int, capture_date: _dt.date) -> list:
    """One row per <parcel>-1.jpg in photos_dir, sorted by numeric parcel id."""
    cap = f"{capture_date.month}/{capture_date.day}/{capture_date.year}"   # M/D/YYYY
    entries = []
    for f in photos_dir.iterdir():
        m = PRIMARY_RE.search(f.name)
        if not m:
            continue
        parid = m.group(1)
        entries.append((int(parid), parid, str(f.resolve())))
    entries.sort(key=lambda e: e[0])
    rows = []
    for _, parid, abspath in entries:
        rows.append([abspath, "", "1", parid, JUR, str(year),
                     "Primary", CARD, "", "", cap, "", "", "", ""])
    return rows


def generate(week_folder: Path, date_str, year: int = None,
             out_path: Path = None, dry_run: bool = False) -> Path:
    """Build the photo-upload CSV for one week folder. Callable from run_weekly.py.

    Returns the output path (or the path it *would* write, on dry-run).
    Raises FileNotFoundError if there is no Photos_New, ValueError if it holds
    no primary photos.
    """
    year = TARGET_TAX_YEAR if year is None else year
    photos_dir = week_folder / "Photos_New"
    if not photos_dir.is_dir():
        raise FileNotFoundError(f"no Photos_New in {week_folder}")

    cap_date = _date_from_folder(week_folder, date_str)
    rows = build_rows(photos_dir, year, cap_date)
    if not rows:
        raise ValueError(f"no <parcel>-1.jpg photos in {photos_dir}")

    if out_path is None:
        out_path = week_folder / f"Photo Upload {cap_date.month}-{cap_date.day}-{cap_date.year}.csv"
    out_path = Path(out_path)

    print(f"  Week folder:  {week_folder.name}")
    print(f"  Photos_New:   {len(rows)} primary photo(s)")
    print(f"  Tax year:     {year}")
    print(f"  Capture date: {cap_date.month}/{cap_date.day}/{cap_date.year}")
    print(f"  Output:       {out_path}")
    if dry_run:
        print("  (dry-run — nothing written)")
        return out_path

    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(HEADER)
        w.writerows(rows)
    print(f"  ✓ Wrote {len(rows)} rows (Taxyr={year})")
    return out_path


def main():
    ap = argparse.ArgumentParser(description="Generate the iasWorld photo-upload CSV for a week.")
    ap.add_argument("--week", default=None, help="week folder name, e.g. 7-13-26 (default: latest)")
    ap.add_argument("--year", type=int, default=TARGET_TAX_YEAR,
                    help=f"tax year for the Taxyr column (default: {TARGET_TAX_YEAR}, from create_future_year)")
    ap.add_argument("--out", default=None, help="output CSV path (default: 'Photo Upload M-D-YYYY.csv' in the week folder)")
    ap.add_argument("--dry-run", action="store_true", help="print a summary, write nothing")
    args = ap.parse_args()

    if args.week:
        week_folder = MLSCAMA_ROOT / args.week
        if not week_folder.is_dir():
            print(f"ERROR: week folder not found: {week_folder}"); sys.exit(1)
        xlsxs = list(week_folder.glob("value_mismatches_*.xlsx"))
        date_str = xlsxs[0].stem.replace("value_mismatches_", "") if xlsxs else None
    else:
        week_folder, date_str = find_latest_week()

    try:
        generate(week_folder, date_str, year=args.year,
                 out_path=Path(args.out) if args.out else None, dry_run=args.dry_run)
    except (FileNotFoundError, ValueError) as e:
        print(f"ERROR: {e}"); sys.exit(1)


if __name__ == "__main__":
    main()
