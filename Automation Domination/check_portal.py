"""
check_portal.py — pre-push sanity checks on the built review_portal.html.

Runs every check in ONBOARDING.md §4 in one go, so the operator does not have to
grep by hand. Exits 0 when everything passes, 1 on any failure.

    python "Automation Domination\\check_portal.py"
"""

import json
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PORTAL_ROOT = Path(__file__).parent.parent
HTML        = PORTAL_ROOT / "review_portal.html"
VERSION     = PORTAL_ROOT / "version.json"
BABEL       = PORTAL_ROOT / "lib" / "babel.min.js"

EXPECTED_MODEL = "claude-sonnet-4-6"


def main() -> int:
    html = HTML.read_text(encoding="utf-8")
    results = []

    def check(name, ok, detail=""):
        results.append(ok)
        print(f"  {'✓' if ok else '✗'} {name}" + (f"  — {detail}" if detail else ""))

    print(f"\nChecking {HTML.name}  ({HTML.stat().st_size / 1e6:.1f} MB)\n")

    check("Anthropic API block present",
          "https://api.anthropic.com/v1/messages" in html)

    models = sorted(set(re.findall(r"claude-[a-z0-9.-]+", html)))
    check(f"Suggest model is {EXPECTED_MODEL}",
          models == [EXPECTED_MODEL], ", ".join(models) or "none found")

    m = re.search(r'generatedAt:\s*"([^"]+)"', html)
    page_ts = m.group(1) if m else None
    file_ts = json.loads(VERSION.read_text(encoding="utf-8")).get("generatedAt")
    check("generatedAt matches version.json",
          page_ts is not None and page_ts == file_ts, f"page={page_ts!r} file={file_ts!r}")

    m = re.search(r"EFF_AGE_BASE_YEAR\s*=\s*(\d+)", html)
    year = int(m.group(1)) if m else None
    check("EFF_AGE_BASE_YEAR is the current year",
          year == date.today().year, f"{year}")

    cdn = re.findall(r"unpkg\.com|cdn\.sheetjs\.com", html)
    check("No CDN library links (libs self-hosted)", not cdn, f"{len(cdn)} found" if cdn else "")

    # sk-ant- legitimately appears as a placeholder/hint; a real key is longer.
    keys = re.findall(r"sk-ant-[A-Za-z0-9_\-]{6,}", html)
    check("No real API key embedded", not keys, f"{len(keys)} found" if keys else "")

    m = re.search(r"IASWORLD_CLOUD_DEFAULT\s*=\s*(true|false)", html)
    check("iasWorld button points at Tyler cloud", bool(m) and m.group(1) == "true",
          m.group(1) if m else "flag not found")

    # Babel parse — a broken script ships a blank navy screen.
    blocks = re.findall(r'<script type="text/babel"[^>]*>(.*?)</script>', html, re.S)
    node = shutil.which("node")
    if not blocks:
        check("Babel script parses", False, "no text/babel block found")
    elif not node:
        print("  - Babel parse skipped (Node.js not installed)")
    else:
        with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False, encoding="utf-8") as f:
            f.write(max(blocks, key=len))
        js = ("const B=require(process.argv[1]);"
              "B.transform(require('fs').readFileSync(process.argv[2],'utf8'),{presets:['react']});")
        r = subprocess.run([node, "-e", js, str(BABEL), f.name], capture_output=True, text=True)
        Path(f.name).unlink(missing_ok=True)
        check("Babel script parses", r.returncode == 0, r.stderr.strip().splitlines()[0] if r.returncode else "")

    ok = all(results)
    print(f"\n  {'PASS — safe to push' if ok else 'FAIL — do not push; fix the ✗ items above'}\n")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
