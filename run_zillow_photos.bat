@echo off
cd /d "C:\Users\jmjeffri\OneDrive - Stark County\Desktop\Portal Builder"

echo ============================================================
echo  Zillow Photo Downloader - Value Mismatches
echo ============================================================
python "ZillowPhotos/download_zillow_photos.py" "MLSvsCAMA/8-24-26/value_mismatches_2026-08-24.xlsx" "MLSvsCAMA/8-24-26/Photos_New"

echo.
echo ============================================================
echo  Zillow Photo Downloader - Perfect Matches
echo ============================================================
python "ZillowPhotos/download_zillow_photos.py" "MLSvsCAMA/8-24-26/perfect_matches_2026-08-24.xlsx" "MLSvsCAMA/8-24-26/Photos_New"

echo.
echo ============================================================
echo  DONE - Close this window when finished
echo ============================================================
pause
