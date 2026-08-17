"""
photo_checks.py — pre-build gate on the week's photo folders.

Runs immediately before the photo-upload CSV is regenerated and the portal is
built, so a bad photo can't reach the live portal or the iasWorld Document
Loader. Called by both run_build.py and run_weekly.py (they duplicate steps 6–7,
so the logic lives here rather than in either one).

Why this exists — three failure modes that repeatedly survived to the build:

  1. A manual photo swap lands in Photos_New but NOT Photos_New_Portal.
     Hit W22, W24, W30, W32, W33 and W34 (2026). Every time, FILENAMES AND FILE
     COUNTS matched in both folders — only a content hash catches it. This is why
     the check compares bytes, not names.
     As of W34 sync_primaries() propagates these automatically before the gate
     runs, so this should no longer reach here. It is still checked: the sync is
     the fix, this is the proof it worked. Note the cause was NOT a copy that
     "silently did not take" — no code writes one folder without the other. See
     sync_primaries() for the actual mechanism.
  2. Zillow returns a realtor headshot instead of a house. The tell is one
     identical hash appearing as the lead photo on several parcels: three
     different houses cannot be byte-identical.
  3. A parcel silently has no primary photo. It still shows secondary photos in
     the portal (photoUrls = [primary, -2, -3, -4].filter(Boolean)), but
     make_photo_upload.py keys on <parcel>-1.jpg, so it gets NO iasWorld upload
     row. W32 nearly shipped 5 such parcels unnoticed.

HARD failures (abort the build) are the ones that put wrong data in front of
staff or into iasWorld and that no human is likely to spot. Coverage gaps are
WARNINGS: Zillow legitimately misses parcels some weeks, and blocking the whole
build for that would be worse than shipping a known, reported gap.
"""

import hashlib
import re
import shutil
from collections import defaultdict
from pathlib import Path

PRIMARY_RE = re.compile(r"^(\d+)-1\.jpg$", re.IGNORECASE)
ANY_PHOTO_RE = re.compile(r"^(\d+)-(\d+)\.(jpg|jpeg|png)$", re.IGNORECASE)

# Lead photos below this are almost never a real exterior shot.
MIN_LEAD_BYTES = 12000
# Exterior shots are usually landscape; below this ratio, ask for an eyeball.
MIN_LEAD_ASPECT = 1.2
# Cap per-parcel listings so a big gap can't bury the hard errors below it.
MAX_LISTED = 15


class PhotoCheckResult:
    def __init__(self):
        self.errors = []    # blocking — build must not proceed
        self.warnings = []  # reported loudly, build continues
        self.stats = {}

    @property
    def ok(self):
        return not self.errors


def _md5(path):
    return hashlib.md5(path.read_bytes()).hexdigest()


def _photos_by_parcel(folder):
    out = defaultdict(list)
    if not folder.is_dir():
        return out
    for f in sorted(folder.iterdir()):
        m = ANY_PHOTO_RE.match(f.name)
        if m:
            out[m.group(1)].append(f)
    return out


def _expected_parcels(mismatches_xlsx, perfects_xlsx):
    """{parcel_id: (mls, address)} across both workbooks. Empty dict if unreadable."""
    try:
        import pandas as pd
    except ImportError:
        return {}
    out = {}
    for path in (mismatches_xlsx, perfects_xlsx):
        if not (path and Path(path).exists()):
            continue
        try:
            df = pd.read_excel(path)
        except Exception:
            continue
        if "Parcel_ID" not in df.columns:
            continue
        for _, r in df.iterrows():
            pid = str(r["Parcel_ID"]).strip()
            if pid and pid not in out:
                out[pid] = (r.get("Listing_Number", ""),
                            f"{r.get('Address', '')}, {r.get('City', '')}".strip(", "))
    return out


def sync_primaries(week_folder, dry_run=False):
    """Propagate a manual <parcel>-1.jpg fix between Photos_New and
    Photos_New_Portal so a swap cannot land in only one folder.

    Why this exists — the failure it removes was misdiagnosed for six weeks
    (W22, W24, W30, W32, W33, W34 of 2026) as a copy that "silently does not
    take". There is no such copy. NO code path writes one folder without the
    other, so nothing existed to fail: download_zillow_photos.py writes photo 1
    to Photos_New_Portal and DERIVES Photos_New from it via shutil.copy2. The
    real data flow is Portal -> Photos_New, the reverse of the mental model in
    the docs. A manual fix was therefore propagated only by the operator
    remembering to write the same file into two directories, with nothing but a
    printed reminder enforcing it — and Photos_New, holding only -1.jpg files,
    looks like "the primaries folder" where a replacement primary belongs.
    Six-for-six is a process gap, not a filesystem gremlin.

    The two folders are required to hold byte-identical primaries (check() hard-
    fails otherwise), so syncing them is enforcing an existing invariant, not a
    policy choice. Direction is decided by mtime — whichever copy the operator
    touched most recently wins — which covers both real cases: saving into
    Photos_New only (the usual one) and promoting a secondary inside
    Photos_New_Portal only (which otherwise leaves the parcel with no iasWorld
    upload row, since make_photo_upload.py keys on Photos_New/<parcel>-1.jpg).
    A fresh download leaves both mtimes equal because copy2 preserves them, so
    an untouched week syncs nothing.

    Only -1.jpg is synced. Secondaries live in Photos_New_Portal by design.
    """
    week_folder = Path(week_folder)
    new_dir = week_folder / "Photos_New"
    portal_dir = week_folder / "Photos_New_Portal"
    actions = []

    if not (new_dir.is_dir() and portal_dir.is_dir()):
        return actions

    names = {f.name for d in (new_dir, portal_dir) if d.is_dir()
             for f in d.iterdir() if PRIMARY_RE.match(f.name)}

    for name in sorted(names):
        a, b = new_dir / name, portal_dir / name
        if a.exists() and b.exists():
            if _md5(a) == _md5(b):
                continue
            # Both present but different — the newer one is the operator's fix.
            src, dst = (a, b) if a.stat().st_mtime >= b.stat().st_mtime else (b, a)
            why = "differed"
        elif a.exists():
            src, dst = a, b
            why = "missing from Photos_New_Portal"
        else:
            src, dst = b, a
            why = "missing from Photos_New"

        arrow = ("Photos_New -> Portal" if src.parent == new_dir
                 else "Portal -> Photos_New")
        actions.append((name, why, arrow))
        if not dry_run:
            shutil.copy2(src, dst)

    return actions


def report_sync(actions, dry_run=False):
    """Print what sync_primaries() did. Silence means the folders agreed."""
    if not actions:
        print("  ✓ Primaries already in sync (no manual swap left behind)")
        return
    verb = "would sync" if dry_run else "synced"
    print(f"  ⚠  {verb} {len(actions)} primary photo(s) — a manual swap "
          f"landed in only one folder:")
    for name, why, arrow in actions:
        print(f"      {name:<18} {why:<32} [{arrow}]")


def check(week_folder, mismatches_xlsx=None, perfects_xlsx=None):
    """Inspect the week's photo folders. Returns PhotoCheckResult (never raises)."""
    week_folder = Path(week_folder)
    new_dir = week_folder / "Photos_New"
    portal_dir = week_folder / "Photos_New_Portal"
    res = PhotoCheckResult()

    if not portal_dir.is_dir():
        res.errors.append(f"Photos_New_Portal does not exist: {portal_dir}")
        return res
    if not new_dir.is_dir():
        res.errors.append(f"Photos_New does not exist: {new_dir}")
        return res

    leads = sorted((f for f in new_dir.iterdir() if PRIMARY_RE.match(f.name)),
                   key=lambda f: f.name)
    if not leads:
        res.errors.append(f"no <parcel>-1.jpg files in {new_dir}")
        return res

    portal_by_parcel = _photos_by_parcel(portal_dir)
    expected = _expected_parcels(mismatches_xlsx, perfects_xlsx)

    # ── HARD: lead photo must exist in the portal folder, byte-identical ──────
    absent, differs = [], []
    lead_hashes = {}
    for f in leads:
        pid = PRIMARY_RE.match(f.name).group(1)
        counterpart = portal_dir / f.name
        if not counterpart.exists():
            absent.append(pid)
            continue
        h_new = _md5(f)
        lead_hashes[pid] = h_new
        if h_new != _md5(counterpart):
            differs.append(pid)

    if absent:
        res.errors.append(
            f"{len(absent)} lead photo(s) missing from Photos_New_Portal "
            f"(portal would show a stale or absent photo): {', '.join(sorted(absent))}")
    if differs:
        res.errors.append(
            f"{len(differs)} lead photo(s) DIFFER between Photos_New and "
            f"Photos_New_Portal — a manual swap only landed in one folder: "
            f"{', '.join(sorted(differs))}")

    # ── HARD: same lead photo on multiple parcels = headshot/placeholder ──────
    by_hash = defaultdict(list)
    for pid, h in lead_hashes.items():
        by_hash[h].append(pid)
    dupes = {h: pids for h, pids in by_hash.items() if len(pids) > 1}
    for h, pids in dupes.items():
        size = (new_dir / f"{pids[0]}-1.jpg").stat().st_size
        res.errors.append(
            f"identical lead photo on {len(pids)} parcels ({size:,} bytes, "
            f"md5 {h[:10]}) — almost certainly a realtor headshot or placeholder: "
            f"{', '.join(sorted(pids))}")

    # ── WARN: parcels with no lead photo → no iasWorld upload row ─────────────
    have_lead = {PRIMARY_RE.match(f.name).group(1) for f in leads}
    if expected:
        no_lead = sorted(set(expected) - have_lead)
        if no_lead:
            lines = []
            for pid in no_lead[:MAX_LISTED]:
                mls, addr = expected[pid]
                extra = len(portal_by_parcel.get(pid, []))
                lines.append(f"      {pid:<10} MLS {mls}  {addr}"
                             + (f"   ({extra} secondary photo(s) — shows in portal, "
                                f"no upload row)" if extra else "   (NO photo at all)"))
            if len(no_lead) > MAX_LISTED:
                lines.append(f"      ... and {len(no_lead) - MAX_LISTED} more "
                             f"(full list: {', '.join(no_lead[MAX_LISTED:])})")
            res.warnings.append(
                f"{len(no_lead)} of {len(expected)} parcels have NO primary photo, so "
                f"they get NO iasWorld photo-upload row:\n" + "\n".join(lines))
        res.stats["expected"] = len(expected)
        res.stats["covered"] = len(expected) - len(no_lead)

    # ── WARN: lead photo geometry looks unlike an exterior shot ──────────────
    try:
        from PIL import Image
        suspects = []
        for f in leads:
            pid = PRIMARY_RE.match(f.name).group(1)
            size = f.stat().st_size
            try:
                with Image.open(f) as im:
                    w, h = im.size
            except Exception:
                continue
            if size < MIN_LEAD_BYTES or (h and w / h < MIN_LEAD_ASPECT):
                suspects.append(f"{pid} ({w}x{h}, {size:,} bytes)")
        if suspects:
            res.warnings.append(
                f"{len(suspects)} lead photo(s) are portrait/square or unusually "
                f"small — eyeball these; phone photos of real houses are often "
                f"portrait, so this is a hint, not proof:\n      "
                + "\n      ".join(suspects))
    except ImportError:
        res.warnings.append("Pillow not installed — skipped lead-photo geometry check")

    res.stats["leads"] = len(leads)
    res.stats["portal_files"] = sum(len(v) for v in portal_by_parcel.values())
    return res


def report(res, strict=False):
    """Print the result. Returns True if the build may proceed."""
    s = res.stats
    if s:
        print(f"  Lead photos: {s.get('leads', 0)}"
              + (f"  |  parcels covered: {s.get('covered', '?')}/{s.get('expected', '?')}"
                 if "expected" in s else "")
              + f"  |  portal files: {s.get('portal_files', 0)}")

    for w in res.warnings:
        print(f"  ⚠  {w}")
    for e in res.errors:
        print(f"  ✗  {e}")

    if res.errors:
        print("\n  " + "=" * 66)
        print("  BUILD BLOCKED — fix the photos above, then re-run.")
        print("  Sync fix:  copy Photos_New\\<parcel>-1.jpg over "
              "Photos_New_Portal\\<parcel>-1.jpg")
        print("  Override:  re-run with --skip-photo-checks (not recommended)")
        print("  " + "=" * 66)
        return False

    if strict and res.warnings:
        print("\n  BUILD BLOCKED — warnings present and --strict-photos was set.")
        return False

    print("  ✓ Photo checks passed"
          + (" (with warnings above)" if res.warnings else ""))
    return True


def gate(week_folder, mismatches_xlsx=None, perfects_xlsx=None,
         skip=False, strict=False):
    """Convenience wrapper: check + report. Returns True if build may proceed."""
    if skip:
        print("  ⚠  photo checks SKIPPED (--skip-photo-checks)")
        return True
    return report(check(week_folder, mismatches_xlsx, perfects_xlsx), strict=strict)
