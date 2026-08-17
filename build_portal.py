"""
build_portal.py
───────────────────────────────────────────────────────────────
Stark County Auditor — MLS/CAMA Review Portal
Monday deploy script

Usage (double-click or run in terminal):
    python build_portal.py

Or with explicit paths:
    python build_portal.py --mismatches "C:/path/value_mismatches.xlsx"
                           --perfects   "C:/path/perfect_matches.xlsx"
                           --photos     "C:/path/to/photos/"
                           --output     "review_portal.html"

Requirements:
    pip install openpyxl pillow
"""

import argparse
import base64
import io
import json
import os
import re
import sys
from pathlib import Path

# Force UTF-8 output on Windows (prevents cp1252 errors for checkmarks, arrows, etc.)
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# ── Try to import dependencies with helpful error messages ─────────────────────
try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl not installed. Run:  pip install openpyxl")
    sys.exit(1)

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("WARNING: Pillow not installed — photos will be embedded at full size.")
    print("         Run:  pip install pillow   for smaller output files.\n")

# ── Configuration ──────────────────────────────────────────────────────────────
PHOTO_MAX_WIDTH  = 400   # px — resize photos to this width max (keeps file size small)
PHOTO_MAX_HEIGHT = 300   # px
PHOTO_QUALITY    = 72    # JPEG quality for embedded photos
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif"}

# ── API Key (set this once — paste your Anthropic API key below) ───────────────
# Get your key from: https://console.anthropic.com → API Keys → Create Key
# Leave as empty string "" if you don't want to embed it (staff will set manually)
ANTHROPIC_API_KEY = ""   # ← PASTE YOUR KEY HERE e.g. "sk-ant-api03-..."

# ── GitHub Pages base URL (for loading archive JSON files) ─────────────────────
# This is the raw base of your GitHub Pages site. Archive files are fetched from:
#   GITHUB_PAGES_BASE + "/archive/YYYY-WNN.json"
GITHUB_PAGES_BASE = "https://starkcountyohio.github.io/cama-review-portal"

# ── Google Apps Script Web App URL (shared status backend) ────────────────────
# Already deployed — do not change unless you redeploy the Apps Script
SHARED_API_URL = "https://script.google.com/macros/s/AKfycbxv8Y30ThaDcxzCymkMxtMm3gTsmAnisF0yhZDI2au2VUNYc9Ypwkevat03tI1KimeP4w/exec"


# ── Helpers ────────────────────────────────────────────────────────────────────

def xlsx_to_json(path: Path) -> list[dict]:
    """Read an xlsx file and return list of row dicts."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(rows[0])]
    result = []
    for row in rows[1:]:
        if all(v is None for v in row):
            continue
        result.append({headers[i]: (row[i] if i < len(row) else None) for i in range(len(headers))})
    wb.close()
    return result


def photo_key_from_filename(name: str) -> str:
    """
    Maps filename to photoMap key used in the portal.
    '102774-1.jpg' → '102774'   (primary — key = parcel ID, backward-compatible)
    '102774-2.jpg' → '102774-2' (extra photo — key includes photo number)
    '102774-3.jpg' → '102774-3'
    """
    stem = Path(name).stem  # e.g. '102774-1'
    m = re.match(r'^(\d+)-(\d+)$', stem)
    if m:
        parcel_id, photo_num = m.group(1), int(m.group(2))
        return parcel_id if photo_num == 1 else f"{parcel_id}-{photo_num}"
    # Fallback: extract leading digits
    m2 = re.match(r'^(\d+)', stem)
    return m2.group(1) if m2 else stem


def image_to_base64(path: Path) -> str | None:
    """Read an image, resize it, return base64 data URL."""
    try:
        if PIL_AVAILABLE:
            with Image.open(path) as img:
                # Convert to RGB (handles PNG with alpha, etc.)
                if img.mode not in ("RGB", "L"):
                    img = img.convert("RGB")
                # Resize maintaining aspect ratio
                img.thumbnail((PHOTO_MAX_WIDTH, PHOTO_MAX_HEIGHT), Image.LANCZOS)
                buf = io.BytesIO()
                img.save(buf, format="JPEG", quality=PHOTO_QUALITY, optimize=True)
                b64 = base64.b64encode(buf.getvalue()).decode()
                return f"data:image/jpeg;base64,{b64}"
        else:
            # No Pillow — embed raw (larger file)
            with open(path, "rb") as f:
                raw = f.read()
            ext = path.suffix.lower().lstrip(".")
            mime = "image/jpeg" if ext in ("jpg", "jpeg") else f"image/{ext}"
            b64 = base64.b64encode(raw).decode()
            return f"data:{mime};base64,{b64}"
    except Exception as e:
        print(f"  SKIP {path.name}: {e}")
        return None


def load_photos(photo_dir: Path) -> dict[str, str]:
    """Load all images from directory, return {parcelKey: dataUrl}."""
    photo_map = {}
    if not photo_dir or not photo_dir.is_dir():
        return photo_map

    files = [f for f in photo_dir.iterdir()
             if f.is_file() and f.suffix.lower() in IMAGE_EXTENSIONS]
    files.sort()

    print(f"\nLoading {len(files)} photos from {photo_dir} ...")
    for i, f in enumerate(files, 1):
        key = photo_key_from_filename(f.name)
        data_url = image_to_base64(f)
        if data_url:
            photo_map[key] = data_url
        if i % 20 == 0:
            print(f"  {i}/{len(files)} processed...")

    print(f"  Done — {len(photo_map)} photos embedded.")
    return photo_map


def rows_to_safe_json(rows: list[dict]) -> str:
    """Serialize rows to JSON, converting any non-serializable values."""
    def clean(v):
        if v is None:
            return ""
        if isinstance(v, (int, float, bool)):
            return v
        return str(v)

    cleaned = [{k: clean(v) for k, v in row.items()} for row in rows]
    return json.dumps(cleaned)


def get_iso_week_key(date=None) -> str:
    """Return ISO week key like '2026-W11' for the given date (default today)."""
    import datetime
    d = date or datetime.date.today()
    iso = d.isocalendar()
    return f"{iso[0]}-W{str(iso[1]).zfill(2)}"


def save_archive(mismatches_rows: list, perfects_rows: list,
                 week_key: str, week_label: str, output_dir: Path) -> Path:
    """
    Save this week's parcel data to archive/YYYY-WNN.json next to the portal.
    Returns the path written. Photos are NOT archived (too large for GitHub).
    """
    import datetime
    archive_dir = output_dir / "archive"
    archive_dir.mkdir(exist_ok=True)

    def clean(v):
        if v is None:
            return ""
        if isinstance(v, (int, float, bool)):
            return v
        return str(v)

    # Count unique parcels (some rows may share a parcel ID)
    parcel_ids = set()
    for row in mismatches_rows:
        pid = row.get("Parcel_ID") or row.get("PARID") or ""
        if pid:
            parcel_ids.add(str(pid).strip())
    for row in perfects_rows:
        pid = row.get("Parcel_ID") or row.get("PARID") or ""
        if pid:
            parcel_ids.add(str(pid).strip())
    
    payload = {
        "weekKey":      week_key,
        "weekLabel":    week_label,
        "savedAt":      datetime.datetime.now().strftime('%B %d, %Y at %I:%M %p'),
        "totalParcels": len(parcel_ids),  # Authoritative parcel count for this week
        "mismatches":   [{k: clean(v) for k, v in row.items()} for row in mismatches_rows],
        "perfects":     [{k: clean(v) for k, v in row.items()} for row in perfects_rows],
    }

    archive_path = archive_dir / f"{week_key}.json"
    with open(archive_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, separators=(',', ':'))  # compact — smaller file

    size_kb = archive_path.stat().st_size / 1024
    print(f"  Archive:    {archive_path.name}  ({size_kb:,.0f} KB)")
    return archive_path


# ── Build HTML ─────────────────────────────────────────────────────────────────

def build_html(mismatches_rows: list, perfects_rows: list,
               photo_map: dict, week_label: str, generated_at: str,
               api_key: str = "", github_pages_base: str = "", shared_api_url: str = "") -> str:
    """Inject data into the portal HTML template."""

    mismatches_json = rows_to_safe_json(mismatches_rows)
    perfects_json   = rows_to_safe_json(perfects_rows)
    photo_json      = json.dumps(photo_map)

    # Sizes for reporting
    mm_kb  = len(mismatches_json) / 1024
    pm_kb  = len(perfects_json)   / 1024
    ph_kb  = len(photo_json)      / 1024

    print(f"\n  Mismatch data:  {mm_kb:,.0f} KB  ({len(mismatches_rows)} rows)")
    print(f"  Perfect data:   {pm_kb:,.0f} KB  ({len(perfects_rows)} rows)")
    print(f"  Photo data:     {ph_kb:,.0f} KB  ({len(photo_map)} photos)")

    # The injected data block — goes into the HTML before the app script
    data_block = f"""
    // ── PRE-LOADED WEEKLY DATA (injected by build_portal.py) ─────────────────
    window.PRELOADED_DATA = {{
      weekLabel:        {json.dumps(week_label)},
      mismatchRows:     {mismatches_json},
      perfectRows:      {perfects_json},
      photoMap:         {photo_json},
      generatedAt:      {json.dumps(generated_at)},
    }};
    // ── PRE-LOADED API KEY (set by admin in build_portal.py) ──────────────────
    window.PRELOADED_API_KEY = {json.dumps(api_key)};
    // ── GitHub Pages base URL (for archive JSON fetching) ─────────────────────
    window.GITHUB_PAGES_BASE = {json.dumps(github_pages_base)};
    // ── Google Apps Script URL (shared status backend) ────────────────────────
    window.INJECTED_SHARED_API_URL = {json.dumps(shared_api_url)};
    """

    # Read the template HTML
    template_path = Path(__file__).parent / "review_portal_template.html"
    if not template_path.exists():
        print(f"\nERROR: Template not found at {template_path}")
        print("Make sure review_portal_template.html is in the same folder as this script.")
        sys.exit(1)

    with open(template_path, "r", encoding="utf-8") as f:
        html = f.read()

    # Inject data block right before the </head> tag
    if "<!-- PRELOADED_DATA_INJECTION -->" in html:
        html = html.replace("<!-- PRELOADED_DATA_INJECTION -->",
                            f"<script>{data_block}</script>")
    else:
        html = html.replace("</head>",
                            f"<script>{data_block}</script>\n</head>")

    return html


# ── Interactive file picker (Windows) ─────────────────────────────────────────

def pick_file_windows(title: str, filetypes: list) -> Path | None:
    """Open a Windows file dialog. Returns Path or None if cancelled."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        root.destroy()
        return Path(path) if path else None
    except Exception:
        return None


def pick_folder_windows(title: str) -> Path | None:
    """Open a Windows folder dialog. Returns Path or None if cancelled."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        path = filedialog.askdirectory(title=title)
        root.destroy()
        return Path(path) if path else None
    except Exception:
        return None


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Build the weekly MLS/CAMA Review Portal HTML file."
    )
    parser.add_argument("--mismatches", help="Path to value_mismatches .xlsx file")
    parser.add_argument("--perfects",   help="Path to perfect_matches .xlsx file")
    parser.add_argument("--photos",     help="Path to folder containing MLS photos")
    parser.add_argument("--output",     help="Output HTML filename (default: review_portal.html)",
                        default="review_portal.html")
    args = parser.parse_args()

    print("=" * 60)
    print("  Stark County Auditor — Portal Builder")
    print("=" * 60)

    # ── Auto-detect latest week folder in MLSvsCAMA/ ──────────────────────────
    def find_latest_week():
        """Return (mismatches, perfects, photos_dir) from the most recent MLSvsCAMA subfolder."""
        mlscama = Path(__file__).parent / "MLSvsCAMA"
        if not mlscama.is_dir():
            return None, None, None
        # Sort subfolders by modification time, newest first
        folders = sorted(
            [d for d in mlscama.iterdir() if d.is_dir()],
            key=lambda d: d.stat().st_mtime,
            reverse=True
        )
        for folder in folders:
            mm  = next(folder.glob("value_mismatches_*.xlsx"),  None)
            pm  = next(folder.glob("perfect_matches_*.xlsx"),   None)
            ph  = folder / "Photos_New_Portal"
            if mm and pm:
                return mm, pm, (ph if ph.is_dir() else None)
        return None, None, None

    # ── Resolve mismatches file ───────────────────────────────────────────────
    mismatches_path = Path(args.mismatches) if args.mismatches else None
    if not mismatches_path or not mismatches_path.exists():
        auto_mm, auto_pm, auto_ph = find_latest_week()
        if auto_mm:
            mismatches_path = auto_mm
            perfects_path   = auto_pm
            photos_path     = auto_ph
            print(f"\n  Auto-detected week folder: {auto_mm.parent.name}")
            print(f"  Mismatches: {mismatches_path.name}")
            print(f"  Perfects:   {perfects_path.name}")
            print(f"  Photos:     {photos_path.name if photos_path else '(none)'}")
        else:
            print("\nStep 1/3: Select VALUE MISMATCHES Excel file...")
            mismatches_path = pick_file_windows(
                "Select Value Mismatches Excel file",
                [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            if not mismatches_path:
                print("Cancelled.")
                sys.exit(0)
            print(f"  Mismatches: {mismatches_path.name}")
            perfects_path = None
            photos_path   = None
    else:
        perfects_path = None
        photos_path   = None

    # ── Resolve perfects file (only if not already set by auto-detect) ────────
    if not perfects_path:
        perfects_path = Path(args.perfects) if args.perfects else None
    if not perfects_path or not perfects_path.exists():
        print("\nStep 2/3: Select PERFECT MATCHES Excel file...")
        perfects_path = pick_file_windows(
            "Select Perfect Matches Excel file",
            [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not perfects_path:
            print("Cancelled.")
            sys.exit(0)
        print(f"  Perfects:   {perfects_path.name}")

    # ── Resolve photos folder (only if not already set by auto-detect) ────────
    if photos_path is None and args.photos:
        photos_path = Path(args.photos)
    if photos_path is None:
        # Already handled above — no dialog needed when auto-detect ran
        if not (mismatches_path and perfects_path and
                mismatches_path.parent == perfects_path.parent):
            print("\nStep 3/3: Select PHOTOS folder (or press Cancel to skip)...")
            photos_path = pick_folder_windows("Select MLS Photos folder (Cancel to skip)")
            if photos_path:
                print(f"  Photos:     {photos_path}")
            else:
                print("  Photos:     (skipped)")

    # ── Read Excel files ──────────────────────────────────────────────────────
    print("\nReading Excel files...")
    try:
        mismatches_rows = xlsx_to_json(mismatches_path)
        print(f"  Mismatches: {len(mismatches_rows)} rows")
    except Exception as e:
        print(f"ERROR reading mismatches file: {e}")
        sys.exit(1)

    try:
        perfects_rows = xlsx_to_json(perfects_path)
        print(f"  Perfects:   {len(perfects_rows)} rows")
    except Exception as e:
        print(f"ERROR reading perfects file: {e}")
        sys.exit(1)

    # ── Load photos ───────────────────────────────────────────────────────────
    photo_map = load_photos(photos_path) if photos_path else {}

    # ── Determine week label ──────────────────────────────────────────────────
    import datetime
    today = datetime.date.today()
    # Find Monday of current week
    monday = today - datetime.timedelta(days=today.weekday())
    week_label = f"Week of {monday.strftime('%B %d, %Y')}"
    print(f"\nWeek label: {week_label}")

    # ── Get ISO week key ─────────────────────────────────────────────────────
    week_key = get_iso_week_key()

    # ── Build HTML ────────────────────────────────────────────────────────────
    print("\nBuilding HTML...")
    # One build marker shared by the page and version.json so staff browsers can
    # detect when a newer review_portal.html has been published (cache-busted).
    generated_at = datetime.datetime.now().strftime('%B %d, %Y at %I:%M %p')
    html = build_html(mismatches_rows, perfects_rows, photo_map, week_label, generated_at, ANTHROPIC_API_KEY, GITHUB_PAGES_BASE, SHARED_API_URL)

    # ── Write output ──────────────────────────────────────────────────────────
    output_path = Path(args.output)
    if not output_path.is_absolute():
        output_path = Path(__file__).parent / output_path

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    # ── Write version.json (build marker for the reload-prompt) ──────────────
    version_path = output_path.parent / "version.json"
    with open(version_path, "w", encoding="utf-8") as f:
        json.dump({"generatedAt": generated_at,
                   "weekLabel": week_label,
                   "weekKey": week_key}, f)

    # ── Save archive JSON (parcel data for this week, no photos) ─────────────
    print("\nSaving archive...")
    archive_path = save_archive(mismatches_rows, perfects_rows,
                                week_key, week_label, output_path.parent)

    size_mb = output_path.stat().st_size / (1024 * 1024)
    key_status = "✓ Embedded" if ANTHROPIC_API_KEY else "✗ Not set (staff must enter manually)"
    print(f"\n{'=' * 60}")
    print(f"  OUTPUT:  {output_path}")
    print(f"  SIZE:    {size_mb:.1f} MB")
    print(f"  API KEY: {key_status}")
    print(f"{'=' * 60}")
    print("\nNext steps:")
    print("  1. Go to github.com/Starkcountyohio/cama-review-portal")
    print(f"  2. Upload  review_portal.html  (replaces last week\'s portal)")
    print(f"  3. Upload  archive/{week_key}.json  (new — keeps this week\'s parcel data)")
    print(f"  4. Upload  version.json  (build marker — powers the auto reload-prompt)")
    print("     → In GitHub: drag all three files into the repo, commit directly to main")
    print("  5. Staff get a 'Reload now' prompt within 5 min — or on next refresh")
    print()
    print("  TIP: The archive/ folder grows by one .json file each Monday.")
    print("       Never delete old archive files — that\'s how past weeks stay viewable.")
    print()

    # ── Auto-open the file locally to verify ─────────────────────────────────
    try:
        import webbrowser
        ans = input("Open in browser to verify? (y/n): ").strip().lower()
        if ans == 'y':
            webbrowser.open(output_path.as_uri())
    except Exception:
        pass

    input("\nPress Enter to close...")


if __name__ == "__main__":
    main()
