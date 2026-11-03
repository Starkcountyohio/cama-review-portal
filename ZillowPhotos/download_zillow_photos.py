"""
Zillow Photo Downloader - Uses Playwright (works with Python 3.14)
Downloads property photos from Zillow based on Excel address list

Usage: python download_zillow_photos.py "perfect_matches.xlsx" [photos_dir]

  photos_dir    Folder for exterior photos (iasWorld CAMA upload). Default: cama_photos
                A sibling folder "{photos_dir}_Portal" is created automatically for all photos
                (up to 4 per parcel) used by the Review Portal build.

Install:
    pip install playwright pandas openpyxl
    playwright install chromium

Output:
  {photos_dir}/             ParcelID-1.jpg only  → iasWorld Document Loader
  {photos_dir}_Portal/      ParcelID-1 through ParcelID-4.jpg  → portal build
  Photo Upload M-D-YYYY.csv auto-generated next to {photos_dir}
"""

from playwright.sync_api import sync_playwright
import csv
import io
import json
import os
import re
import shutil
import sys
import time
import pandas as pd
from datetime import datetime
from pathlib import Path
from urllib.parse import quote
from PIL import Image
import random

# Photos saved per parcel: the listing's first N, as -1 … -N.
PHOTOS_PER_PARCEL = 4

# Seconds to wait between parcels. Zillow's page loads add ~15-25 s on top, so
# 35-50 s gives ~1 minute per parcel in total. 60-75 s measured ~90 s per parcel
# (2026-11-03 test, 1 CAPTCHA in 5); 30-60 s drew frequent CAPTCHAs (W45).
PARCEL_DELAY = (35, 50)

# Force UTF-8 output on Windows (prevents cp1252 errors for arrows, checkmarks, etc.)
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# ── Proxy configuration ────────────────────────────────────────────────────────
# Paste your Smartproxy (or other rotating residential proxy) URL here.
# Format:  "http://USERNAME:PASSWORD@gate.smartproxy.com:10000"
# Leave as "" to run without a proxy (you may hit Zillow CAPTCHAs faster).
# Can also be set via the ZILLOW_PROXY environment variable or --proxy flag.
ZILLOW_PROXY = ""   # ← PASTE YOUR PROXY URL HERE


class ZillowDownloader:
    """Download Zillow photos using Playwright"""

    def __init__(self, output_dir="cama_photos", proxy=None):
        photos_path = Path(output_dir)
        # CAMA folder: exterior -1.jpg only → iasWorld Document Loader
        self.photos_dir = str(photos_path)
        # Portal folder: all photos up to -4.jpg → Review Portal build
        self.portal_dir = str(photos_path.parent / (photos_path.name + "_Portal"))

        # Optional proxy: "http://user:pass@host:port" or "socks5://host:port"
        # Rotating residential proxy (Bright Data, Smartproxy, IPRoyal, etc.)
        # prevents Zillow IP-based rate limiting. Each connection exits from a
        # different residential IP automatically — no extra rotation code needed.
        # Priority: explicit arg → module constant → environment variable
        self.proxy = proxy or ZILLOW_PROXY or os.environ.get("ZILLOW_PROXY", "")

        self.browser = None
        self.page = None
        self.context = None
        self.playwright = None

        os.makedirs(self.photos_dir, exist_ok=True)
        os.makedirs(self.portal_dir, exist_ok=True)

    def _proxy_config(self):
        """Return Playwright proxy dict, or None if no proxy configured."""
        if not self.proxy:
            return None
        # Support "user:pass@host:port" embedded in the URL, or plain "host:port"
        server = self.proxy if "://" in self.proxy else f"http://{self.proxy}"
        return {"server": server}

    def _new_context(self):
        """Create (or recreate) a browser context, applying proxy if configured."""
        if self.context:
            try:
                self.context.close()
            except Exception:
                pass
        proxy_cfg = self._proxy_config()
        self.context = self.browser.new_context(
            proxy=proxy_cfg,
            ignore_https_errors=True,
            viewport={"width": 1280, "height": 800}
        )
        self.page = self.context.new_page()

    def start_browser(self):
        """Start Chrome via Playwright"""
        print("\nStarting Chrome (Playwright)...")
        if self.proxy:
            # Show host only (hide credentials in the log)
            try:
                from urllib.parse import urlparse
                host = urlparse(self.proxy if "://" in self.proxy
                                else f"http://{self.proxy}").hostname
            except Exception:
                host = self.proxy.split("@")[-1].split(":")[0]
            print(f"  Proxy: {host}  (rotating residential)")
        try:
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(
                headless=False,
                args=[
                    '--start-maximized',
                    '--disable-blink-features=AutomationControlled',
                    '--ignore-certificate-errors',
                ]
            )
            self._new_context()
            print("✓ Chrome started")
            return True
        except Exception as e:
            print(f"✗ Error starting browser: {e}")
            print("\n🔧 Fix:")
            print("  py -m pip install playwright")
            print("  py -m playwright install chromium")
            return False

    def _is_challenge_page(self):
        """Return True if the current page looks like a bot-challenge page."""
        try:
            url = self.page.url.lower()
            if '/homedetails/' in url or '/b/' in url:
                return False  # Normal property or search-results page
            if 'captcha' in url or 'challenge' in url:
                return True
            # Check page title for access-denied / challenge phrases
            title = (self.page.title() or '').lower()
            if any(k in title for k in ('access to this page', 'denied', 'blocked', 'captcha', 'challenge', 'verify')):
                return True
            # Check visible text for "Press & Hold" style challenges
            body = (self.page.inner_text('body') or '').lower()
            if any(k in body for k in ('press & hold', 'press and hold', 'confirm you are a human',
                                        'are you a robot', 'verify you are human')):
                return True
            # Iframe-based recaptcha
            for selector in ('iframe[src*="recaptcha"]', 'iframe[src*="captcha"]',
                             '#captcha', 'iframe[title*="recaptcha"]'):
                if self.page.query_selector(selector):
                    return True
        except Exception:
            pass
        return False

    def check_for_captcha(self):
        """Check if CAPTCHA is present and wait (by polling) until it is solved."""
        try:
            if not self._is_challenge_page():
                return False

            print("\n" + "!" * 60)
            print("⚠️  CAPTCHA DETECTED!")
            print("    Solve the CAPTCHA in the browser window.")
            print("    Script will continue automatically once it is solved.")
            print("!" * 60)
            # Poll every 2 s until the challenge page is gone
            while True:
                time.sleep(2)
                try:
                    if not self._is_challenge_page():
                        print("    ✓ CAPTCHA solved — waiting for page to load...")
                        # Wait for the page to fully load after the challenge clears
                        try:
                            self.page.wait_for_load_state("networkidle", timeout=10000)
                        except Exception:
                            pass
                        time.sleep(2)
                        break
                except Exception:
                    break
            return True

        except Exception:
            pass
        return False

    def search_zillow(self, address, city, state, zipcode):
        """Search Zillow for property"""
        try:
            full_address = f"{address}, {city}, {state} {zipcode}"
            # Encode the address: an unescaped "#" (unit numbers like "#I") starts a
            # URL fragment, so Zillow only saw the street and returned another unit
            # (W45: 290 Lincoln St SW #I landed on APT A).
            url = f"https://www.zillow.com/homes/{quote(full_address, safe=', ')}_rb/"

            print(f"    Searching...")
            self.page.goto(url, wait_until="domcontentloaded", timeout=30000)
            time.sleep(random.uniform(2, 4))

            self.check_for_captcha()

            current_url = self.page.url

            # Only trust Zillow's own redirect to a property page. A results page
            # means Zillow had no exact match; its first link is a guess, and was
            # the source of wrong-house photos.
            if '/homedetails/' in current_url:
                print(f"    → Found")
                return current_url

            print(f"    → Not found (no exact match on Zillow)")
            return None

        except Exception as e:
            print(f"    → Error: {str(e)[:50]}...")
            return None

    @staticmethod
    def _page_property(content):
        """Return the property record Zillow embeds in the page, or None.

        Zillow's JSON-LD block used to carry the whole gallery; since ~2026-10-26
        it holds only the lead photo, which is why W44/W45 saved one photo per
        parcel. The full gallery for THIS property lives in __NEXT_DATA__ ->
        gdpClientCache -> <query> -> property. Scanning the raw page source
        instead (the old fallback) also picks up the "similar homes" carousels —
        other houses' photos, which is where W45's wrong-house photos came from.
        """
        m = re.search(r'<script[^>]+id="__NEXT_DATA__"[^>]*>(.*?)</script>', content, re.DOTALL)
        if not m:
            return None
        try:
            cache = json.loads(m.group(1))["props"]["pageProps"]["componentProps"]["gdpClientCache"]
            if isinstance(cache, str):
                cache = json.loads(cache)
        except (ValueError, KeyError, TypeError):
            return None
        for entry in cache.values():
            prop = entry.get("property") if isinstance(entry, dict) else None
            if isinstance(prop, dict) and ("originalPhotos" in prop or "responsivePhotos" in prop):
                return prop
        return None

    @staticmethod
    def _listing_photo_urls(prop):
        """Full-size listing photo URLs in gallery order, Street View excluded."""
        urls = []
        for photo in prop.get("originalPhotos") or prop.get("responsivePhotos") or []:
            jpegs = (photo.get("mixedSources") or {}).get("jpeg") or []
            best = max(jpegs, key=lambda s: s.get("width") or 0, default=None)
            url = (best or {}).get("url") or photo.get("url")
            # When a listing's photos are taken down, Zillow shows a single
            # Google Street View image in their place (W45: 305815, 3603459).
            if url and "photos.zillowstatic.com" in url:
                urls.append(url)
        return urls

    @staticmethod
    def _address_key(street):
        """(house number, first street word, unit) for comparing two addresses."""
        s = re.sub(r"[.,]", " ", str(street).upper())
        unit = re.search(r"(?:#|\b(?:APT|UNIT|STE|SUITE|LOT|BLDG)\b)\s*([A-Z0-9-]+)", s)
        s = s[:unit.start()] if unit else s
        words = s.split()
        number = words[0] if words and words[0][0].isdigit() else ""
        street_word = words[1] if len(words) > 1 else ""
        return number, street_word, unit.group(1) if unit else ""

    def get_all_photo_urls(self, property_url, address):
        """
        Read the property's own photo gallery from its Zillow page.
        Returns (status, urls): status is "ok", "Listing Removed" (gallery gone,
        Street View only), "Address Mismatch", or "Parse Failed".
        """
        try:
            print(f"    Loading page...")
            if self.page.url.split('#')[0] != property_url.split('#')[0]:
                self.page.goto(property_url, wait_until="domcontentloaded", timeout=30000)
            try:
                self.page.wait_for_load_state("networkidle", timeout=20000)
            except Exception:
                pass
            time.sleep(random.uniform(2, 3))

            self.check_for_captcha()

            # The record is server-rendered, but poll briefly in case the page is
            # still settling after a CAPTCHA.
            prop = None
            deadline = time.time() + 20
            while prop is None and time.time() < deadline:
                prop = self._page_property(self.page.content())
                if prop is None:
                    time.sleep(1.5)
            if prop is None:
                print(f"    → Could not read Zillow's property data (page layout changed?)")
                return ("Parse Failed", [])

            theirs = prop.get("streetAddress") or (prop.get("address") or {}).get("streetAddress", "")
            if self._address_key(theirs) != self._address_key(address):
                print(f"    → Zillow page is for '{theirs}', not '{address}' — skipping")
                return ("Address Mismatch", [])

            urls = self._listing_photo_urls(prop)
            if not urls:
                print(f"    → Listing photos taken down (Street View only) — skipping")
                return ("Listing Removed", [])

            print(f"    → Found {len(urls)} listing photo(s)")
            return ("ok", urls)

        except Exception as e:
            print(f"    → Error: {str(e)[:50]}...")
            return ("Parse Failed", [])

    @staticmethod
    def _looks_like_floor_plan(body):
        """True for a floor-plan drawing: mostly near-white and almost colourless.

        Zillow gives no caption or type for floor plans, so this goes by the
        pixels. Measured 2026-11-03 over 339 photos (W43, W45, a live test):
        the four floor plans scored 0.41-0.92 near-white with saturation <= 0.08
        (0.41 was a plan drawn over an aerial shot); the whitest real photo was
        0.12. Cutoffs sit between: >= 0.30 white AND <= 0.15 saturation.
        """
        try:
            from PIL import ImageStat
            img = Image.open(io.BytesIO(body)).convert("RGB")
            img.thumbnail((200, 200))
            px = list(img.getdata())
            white = sum(1 for r, g, b in px if min(r, g, b) >= 235) / len(px)
            saturation = ImageStat.Stat(img.convert("HSV")).mean[1] / 255
            return white >= 0.30 and saturation <= 0.15
        except Exception:
            return False

    def download(self, photo_url, filepath, check_portrait=False):
        """Download photo using Playwright (bypasses SSL issues)"""
        try:
            response = self.page.request.get(photo_url)

            if response.ok:
                body = response.body()
                # Reject tiny placeholder/error responses (< 5 KB)
                if len(body) <= 5000:
                    print(f"    ⚠️  Image too small ({len(body)} bytes) - skipping")
                    return False
                # Screen photo 1 for realtor headshots.
                #
                # This used to reject any portrait image (h > w) on the theory that
                # headshots are portrait and houses are landscape. That was wrong in
                # both directions, and the cost was invisible: when this fires, -1 is
                # written to NEITHER folder while -2/-3/-4 still reach the portal
                # folder, so the parcel silently loses its primary and its iasWorld
                # upload row (mapping CSV shows Success / Photos = 3).
                #   - It rejected genuine houses. Plenty of Stark County listings lead
                #     with a portrait phone shot: W34's 202027 and 209854 were both
                #     576x768 exteriors, dropped here and then pulled by hand.
                #   - It missed the headshots it exists to catch. W34's 247844 was
                #     240x240 — SQUARE, so h > w was false — and the recurring W30/W32
                #     headshot was the same 240x240 / ~8.8 KB image.
                #
                # The actual signature of the headshot family is small AND near-square.
                # Real exteriors are large whatever their shape, so requiring both
                # conditions keeps portrait houses and still drops the headshots.
                #
                # Thresholds are set from the two closest real data points, which are
                # uncomfortably close together — do not loosen them without checking
                # both:
                #   REJECT  247844 (W34) headshot          240x240,  8,860 bytes
                #   KEEP    4307614 (W33) genuine exterior 320x320, 38,127 bytes
                # A 40 KB floor (the obvious first guess) drops that real house. Hence
                # 20 KB / 300 px: above the headshot, below the house, on both axes.
                # Because the margin is this thin, the mandatory photo-review gate
                # stays the real defense — this only trims the obvious cases.
                if check_portrait:
                    try:
                        img = Image.open(io.BytesIO(body))
                        w, h = img.size
                        small = len(body) < 20000 or max(w, h) < 300
                        near_square = 0.9 <= (w / h) <= 1.1 if h else False
                        if small and near_square:
                            print(f"    ⚠️  Small near-square image ({w}x{h}, "
                                  f"{len(body):,} bytes) — likely agent headshot, skipping")
                            return False
                    except Exception:
                        pass
                if self._looks_like_floor_plan(body):
                    print(f"    ⊘ Floor plan — skipping")
                    return False
                with open(filepath, 'wb') as f:
                    f.write(body)
                return True
            return False
        except Exception as e:
            print(f"    Download error: {str(e)[:50]}")
            if os.path.exists(filepath):
                os.remove(filepath)
            return False

    def generate_cama_csv(self):
        """
        Generate iasWorld Document Loader CSV from the Photos/ folder.
        Reads all -1.jpg files and writes 'Photo Upload M-D-YYYY.csv'
        next to the Photos/ folder (same location as the old manual template output).
        """
        photos_path = Path(self.photos_dir)
        today = datetime.now()
        date_for_filename = f"{today.month}-{today.day}-{today.year}"   # 4-14-2026
        date_for_csv      = f"{today.month}/{today.day}/{today.year}"   # 4/14/2026
        csv_filename      = f"Photo Upload {date_for_filename}.csv"
        csv_path          = photos_path.parent / csv_filename

        fieldnames = [
            'Filename', 'FileSize', 'Rank', 'Parid', 'Jur', 'Taxyr',
            'Photo Category', 'Card', 'Title', 'Notes', 'Photo Capture Date',
            'SubjectXCoord', 'GPS_LAT', 'SubjectYCoord', 'GPS_LONG',
        ]

        # Collect -1.jpg files, sorted numerically by parcel ID
        photo_files = [
            f for f in photos_path.iterdir()
            if f.is_file() and f.suffix.lower() in {'.jpg', '.jpeg'} and f.stem.endswith('-1')
        ]
        photo_files.sort(key=lambda f: int(re.match(r'^(\d+)', f.stem).group(1))
                         if re.match(r'^(\d+)', f.stem) else 0)

        rows = []
        for f in photo_files:
            # Extract parcel ID: "202607-1" → "202607"
            parcel_id = f.stem[:-2]  # strip last two chars ("-1")

            rows.append({
                'Filename':           str(f.resolve()),
                'FileSize':           '',
                'Rank':               1,
                'Parid':              parcel_id,
                'Jur':                '000',
                'Taxyr':              today.year,
                'Photo Category':     'Primary',
                'Card':               1,
                'Title':              '',
                'Notes':              '',
                'Photo Capture Date': date_for_csv,
                'SubjectXCoord':      '',
                'GPS_LAT':            '',
                'SubjectYCoord':      '',
                'GPS_LONG':           '',
            })

        if rows:
            with open(csv_path, 'w', newline='', encoding='utf-8') as fh:
                writer = csv.DictWriter(fh, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)
            print(f"  ✓ Document Loader CSV: {csv_path}  ({len(rows)} photos)")
        else:
            print("  (No -1.jpg files found in Photos/ — CSV not generated)")

        return str(csv_path) if rows else None

    def stop_browser(self):
        """Close browser cleanly"""
        try:
            if self.context:
                self.context.close()
            if self.browser:
                self.browser.close()
            if self.playwright:
                self.playwright.stop()
            print("✓ Browser closed")
        except Exception:
            pass

    def process(self, excel_file):
        """Main processing loop"""
        print("\n" + "=" * 80)
        print("ZILLOW PHOTO DOWNLOADER - Playwright edition (Python 3.14 compatible)")
        print("=" * 80)
        print(f"File:          {excel_file}")
        print(f"CAMA folder:   {self.photos_dir}")
        print(f"               (exterior photo 1 only → iasWorld Document Loader)")
        print(f"Portal folder: {self.portal_dir}")
        print(f"               (listing photos 1-{PHOTOS_PER_PARCEL} → Review Portal)")

        try:
            df = pd.read_excel(excel_file)
            print(f"✓ Loaded {len(df)} rows")
        except Exception as e:
            print(f"✗ Could not read file: {e}")
            return None

        df = df.drop_duplicates('Parcel_ID', keep='first')
        print(f"✓ Processing {len(df)} unique properties")
        print("=" * 80)

        if not self.start_browser():
            return None

        results = []

        try:
            for i, (idx, row) in enumerate(df.iterrows()):
                parcel_id = row['Parcel_ID']
                address   = str(row['Address']).strip()
                city      = str(row['City']).strip()
                state     = 'OH' if pd.isna(row.get('State')) else str(row['State']).strip()
                zipcode   = str(int(row['Zip'])) if not pd.isna(row['Zip']) else ''

                print(f"\n[{i+1}/{len(df)}] Parcel {parcel_id}")
                print(f"  {address}, {city}, {state} {zipcode}")

                cama_file = os.path.join(self.photos_dir, f"{parcel_id}-1.jpg")

                # Skip if exterior photo already downloaded
                if os.path.exists(cama_file):
                    print(f"  ⊘ Already downloaded")
                    results.append({'Parcel_ID': parcel_id, 'Photos': 'Exists', 'Status': 'Exists'})
                    continue

                prop_url = self.search_zillow(address, city, state, zipcode)
                if not prop_url:
                    results.append({'Parcel_ID': parcel_id, 'Photos': 0, 'Status': 'Not Found'})
                    # Delay on misses too — skipping it fires back-to-back searches once
                    # Zillow starts refusing, which keeps the block going (W44: 24 misses in a row)
                    time.sleep(random.uniform(*PARCEL_DELAY))
                    continue

                status, all_urls = self.get_all_photo_urls(prop_url, address)
                if status != "ok":
                    results.append({'Parcel_ID': parcel_id, 'Photos': 0, 'Status': status})
                    time.sleep(random.uniform(*PARCEL_DELAY))
                    continue

                # The listing's first photos, in gallery order, floor plans skipped.
                # Photo 1 is the front exterior (-> CAMA + Portal) and is never
                # replaced by a later photo — promoting an interior to the iasWorld
                # primary is a human call. Photos 2+ go to the Portal only.
                saved = 0
                n = 1
                for pos, url in enumerate(all_urls):
                    if n > PHOTOS_PER_PARCEL:
                        break
                    portal_file = os.path.join(self.portal_dir, f"{parcel_id}-{n}.jpg")
                    ok = self.download(url, portal_file, check_portrait=(n == 1))
                    if ok:
                        saved += 1
                        if n == 1:
                            # Copy exterior photo to CAMA folder as well
                            shutil.copy2(portal_file, cama_file)
                            print(f"  ✓ {parcel_id}-1.jpg  (exterior, CAMA + Portal)")
                        else:
                            print(f"  ✓ {parcel_id}-{n}.jpg  (Portal, listing photo {pos + 1})")
                    if ok or n == 1:
                        n += 1

                if saved > 0:
                    results.append({'Parcel_ID': parcel_id, 'Photos': saved, 'Status': 'Success'})
                else:
                    results.append({'Parcel_ID': parcel_id, 'Photos': 0, 'Status': 'Failed'})

                time.sleep(random.uniform(*PARCEL_DELAY))  # Polite delay between parcels; shorter waits drew CAPTCHAs

        except KeyboardInterrupt:
            print("\n\n⚠️  Interrupted by user")
        except Exception as e:
            print(f"\n\n✗ Unexpected error: {e}")
        finally:
            print("\nClosing browser...")
            self.stop_browser()

        df_results = pd.DataFrame(results) if results else pd.DataFrame(columns=['Parcel_ID', 'Photos', 'Status'])
        success = len(df_results[df_results['Status'].isin(['Success', 'Exists'])]) if not df_results.empty else 0

        print("\n" + "=" * 80)
        print(f"COMPLETE! {success}/{len(df)} photos downloaded")
        print("=" * 80)
        print(f"\nCAMA photos:   {self.photos_dir}/")
        print(f"               (ParcelID-1.jpg — ready for iasWorld Document Loader)")
        print(f"Portal photos: {self.portal_dir}/")
        print(f"               (ParcelID-1 through -4.jpg — pass this folder to build_portal.py)")

        # Auto-generate iasWorld Document Loader CSV
        print("\nGenerating iasWorld Document Loader CSV...")
        self.generate_cama_csv()

        timestamp    = datetime.now().strftime("%Y%m%d_%H%M%S")
        mapping_file = os.path.join(self.photos_dir, f"mapping_{timestamp}.csv")
        df_results.to_csv(mapping_file, index=False)
        print(f"Run log:       {mapping_file}")

        return df_results


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Download Zillow property photos for iasWorld CAMA and the Review Portal.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  py download_zillow_photos.py "perfect_matches.xlsx" Photos
  py download_zillow_photos.py "perfect_matches.xlsx" Photos --proxy "http://user:pass@host:port"
  py download_zillow_photos.py "perfect_matches.xlsx" Photos --proxy "brd.superproxy.io:22225"

Proxy (optional):
  Pass --proxy or set the ZILLOW_PROXY environment variable.
  A rotating residential proxy (Bright Data, Smartproxy, IPRoyal) prevents
  Zillow IP rate-limiting. Each connection exits from a different home IP
  automatically — no extra configuration needed beyond the proxy URL.

  Proxy URL format:
    http://USERNAME:PASSWORD@HOST:PORT     (most services)
    socks5://USERNAME:PASSWORD@HOST:PORT
    HOST:PORT                              (no auth, auto-prefixed with http://)

Output per run:
  {photos_dir}/              ParcelID-1.jpg  (CAMA exterior only)
  {photos_dir}_Portal/       ParcelID-1 through -4.jpg  (portal)
  Photo Upload M-D-YYYY.csv  (iasWorld Document Loader — auto-generated)

Install:
  pip install playwright pandas openpyxl
  playwright install chromium
""",
    )
    parser.add_argument("excel_file",  help="Path to perfect_matches or value_mismatches .xlsx")
    parser.add_argument("photos_dir",  nargs="?", default="cama_photos",
                        help="Folder for CAMA exterior photos (default: cama_photos)")
    parser.add_argument("--proxy",     default="",
                        help="Proxy URL, e.g. http://user:pass@host:port  "
                             "(or set ZILLOW_PROXY env var)")

    args = parser.parse_args()

    if not os.path.exists(args.excel_file):
        print(f"\nError: File not found: {args.excel_file}")
        return

    downloader = ZillowDownloader(args.photos_dir, proxy=args.proxy)
    downloader.process(args.excel_file)


if __name__ == "__main__":
    main()
