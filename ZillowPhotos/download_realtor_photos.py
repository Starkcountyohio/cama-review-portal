"""
Realtor.com Photo Downloader - Uses Playwright (works with Python 3.14)
Downloads property photos from Realtor.com based on Excel address list
Less aggressive bot detection than Zillow

Usage: python download_realtor_photos.py "perfect_matches.xlsx" cama_photos

Install: 
    pip install playwright pandas openpyxl
    playwright install chromium
"""

from playwright.sync_api import sync_playwright
import os
import re
import time
import pandas as pd
from datetime import datetime
import sys
import random
import urllib.parse

# Force UTF-8 output on Windows (prevents cp1252 errors for arrows, checkmarks, etc.)
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')


class RealtorDownloader:
    """Download property photos from Realtor.com using Playwright"""

    def __init__(self, output_dir="cama_photos"):
        self.output_dir = output_dir
        self.browser = None
        self.page = None
        self.context = None
        self.playwright = None

        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

    def start_browser(self):
        """Start Chrome via Playwright"""
        print("\nStarting Chrome (Playwright)...")
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
            self.context = self.browser.new_context(
                ignore_https_errors=True,
                viewport={"width": 1280, "height": 800},
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )
            self.page = self.context.new_page()
            print("✓ Chrome started")
            return True
        except Exception as e:
            print(f"✗ Error starting browser: {e}")
            print("\n🔧 Fix:")
            print("  py -m pip install playwright")
            print("  py -m playwright install chromium")
            return False

    def search_realtor(self, address, city, state, zipcode):
        """Search Realtor.com for property"""
        try:
            # Format address for Realtor.com URL
            full_address = f"{address}, {city}, {state} {zipcode}"
            search_query = urllib.parse.quote(full_address)
            url = f"https://www.realtor.com/realestateandhomes-search/{search_query}"

            print(f"    Searching...")
            self.page.goto(url, wait_until="domcontentloaded", timeout=30000)
            time.sleep(random.uniform(3, 5))

            current_url = self.page.url

            # Check if we landed on a property detail page
            if '/realestateandhomes-detail/' in current_url:
                print(f"    → Found (direct)")
                return current_url

            # Look for property cards/links
            try:
                # Try different selectors for property links
                selectors = [
                    "a[href*='/realestateandhomes-detail/']",
                    "[data-testid='property-card'] a",
                    ".property-card a",
                    "a.card-anchor"
                ]
                
                for selector in selectors:
                    links = self.page.query_selector_all(selector)
                    if links:
                        href = links[0].get_attribute('href')
                        if href:
                            print(f"    → Found")
                            return href if href.startswith('http') else f"https://www.realtor.com{href}"
            except Exception:
                pass

            print(f"    → Not found on Realtor.com")
            return None

        except Exception as e:
            print(f"    → Error: {str(e)[:50]}...")
            return None

    def get_photo(self, property_url):
        """Get photo URL from property page"""
        try:
            print(f"    Loading page...")
            self.page.goto(property_url, wait_until="domcontentloaded", timeout=30000)
            time.sleep(random.uniform(3, 5))

            # Try various selectors for Realtor.com photos
            photo_selectors = [
                "img[data-testid='hero-image']",
                "img[data-testid='property-image']",
                ".photo-gallery img",
                "[data-testid='gallery'] img",
                "img[src*='rdcpix.com']",
                "picture img",
                ".media-container img"
            ]
            
            for selector in photo_selectors:
                try:
                    img = self.page.query_selector(selector)
                    if img:
                        src = img.get_attribute('src')
                        if src and len(src) > 50 and ('rdcpix.com' in src or 'realtor.com' in src):
                            print(f"    → Found photo")
                            # Get larger version if possible
                            src = re.sub(r'-[a-z]\d+\.', '-o.', src)  # Try to get original size
                            return src
                except Exception:
                    continue

            # Fallback: regex scan of page source for image URLs
            content = self.page.content()
            
            # Look for Realtor.com image CDN
            urls = re.findall(r'https://[^"\']+rdcpix\.com[^"\']+\.jpg', content)
            if urls:
                # Get the first large-looking image
                for url in urls:
                    if 'thumb' not in url.lower():
                        print(f"    → Found photo (from source)")
                        return url

            print(f"    → No photo found")
            return None

        except Exception as e:
            print(f"    → Error: {str(e)[:50]}...")
            return None

    def download(self, photo_url, filepath):
        """Download photo using Playwright (bypasses SSL issues)"""
        try:
            response = self.page.request.get(photo_url)
            
            if response.ok:
                body = response.body()
                # Real property photos should be at least 30KB
                if len(body) > 30000:
                    with open(filepath, 'wb') as f:
                        f.write(body)
                    print(f"    ({len(body) // 1024}KB)")
                    return True
                else:
                    print(f"    ⚠️  Image too small ({len(body)} bytes)")
                    return False
            return False
        except Exception as e:
            print(f"    Download error: {str(e)[:50]}")
            if os.path.exists(filepath):
                os.remove(filepath)
            return False

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
        print("REALTOR.COM PHOTO DOWNLOADER - Playwright edition (Python 3.14 compatible)")
        print("=" * 80)
        print(f"File:   {excel_file}")
        print(f"Output: {self.output_dir}")
        print(f"Naming: ParcelID-1.jpg  (e.g. 104503-1.jpg)")

        # Load Excel
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
                address = str(row['Address']).strip()
                city = str(row['City']).strip()
                state = 'OH' if pd.isna(row.get('State')) else str(row['State']).strip()
                zipcode = str(int(row['Zip'])) if not pd.isna(row['Zip']) else ''

                print(f"\n[{i+1}/{len(df)}] Parcel {parcel_id}")
                print(f"  {address}, {city}, {state} {zipcode}")

                filename = f"{parcel_id}-1.jpg"
                filepath = os.path.join(self.output_dir, filename)

                # Skip already downloaded
                if os.path.exists(filepath):
                    print(f"  ⊘ Already downloaded")
                    results.append({'Parcel_ID': parcel_id, 'Filename': filename, 'Status': 'Exists'})
                    continue

                prop_url = self.search_realtor(address, city, state, zipcode)
                if not prop_url:
                    results.append({'Parcel_ID': parcel_id, 'Filename': None, 'Status': 'Not Found'})
                    continue

                photo_url = self.get_photo(prop_url)
                if not photo_url:
                    results.append({'Parcel_ID': parcel_id, 'Filename': None, 'Status': 'No Photo'})
                    continue

                print(f"  Downloading...")
                if self.download(photo_url, filepath):
                    print(f"  ✓ Saved: {filename}")
                    results.append({'Parcel_ID': parcel_id, 'Filename': filename, 'Status': 'Success'})
                else:
                    print(f"  ✗ Download failed")
                    results.append({'Parcel_ID': parcel_id, 'Filename': None, 'Status': 'Failed'})

                # Random delay between properties
                time.sleep(random.uniform(5, 10))

        except KeyboardInterrupt:
            print("\n\n⚠️  Interrupted by user")
        except Exception as e:
            print(f"\n\n✗ Unexpected error: {e}")
        finally:
            print("\nClosing browser...")
            self.stop_browser()

        # Summary
        df_results = pd.DataFrame(results)
        success = len(df_results[df_results['Status'].isin(['Success', 'Exists'])])

        print("\n" + "=" * 80)
        print(f"COMPLETE! {success}/{len(df)} photos")
        print("=" * 80)
        print(f"\nPhotos: {self.output_dir}/  (named ParcelID-1.jpg)")
        print(f"Ready to import into iasWorld!")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        mapping_file = os.path.join(self.output_dir, f"mapping_{timestamp}.csv")
        df_results.to_csv(mapping_file, index=False)
        print(f"Mapping CSV: {mapping_file}")

        return df_results


def main():
    if len(sys.argv) < 2:
        print("\n" + "=" * 80)
        print("REALTOR.COM PHOTO DOWNLOADER")
        print("=" * 80)
        print("\nUsage:   python download_realtor_photos.py <excel_file> [output_dir]")
        print("Example: python download_realtor_photos.py \"perfect_matches.xlsx\"")
        print("\nInstall:")
        print("  pip install playwright pandas openpyxl")
        print("  playwright install chromium")
        print("\nNote: Photos named ParcelID-1.jpg (e.g. 104503-1.jpg)")
        return

    excel_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "cama_photos"

    if not os.path.exists(excel_file):
        print(f"\nError: File not found: {excel_file}")
        return

    downloader = RealtorDownloader(output_dir)
    downloader.process(excel_file)


if __name__ == "__main__":
    main()
