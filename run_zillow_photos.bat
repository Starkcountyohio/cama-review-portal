@echo off
cd /d "%~dp0"

echo ============================================================
echo  Zillow Photo Downloader - Value Mismatches
echo ============================================================
python "ZillowPhotos/download_zillow_photos.py" "MLSvsCAMA/10-26-26/value_mismatches_2026-10-26.xlsx" "MLSvsCAMA/10-26-26/Photos_New"

echo.
echo ============================================================
echo  Zillow Photo Downloader - Perfect Matches
echo ============================================================
python "ZillowPhotos/download_zillow_photos.py" "MLSvsCAMA/10-26-26/perfect_matches_2026-10-26.xlsx" "MLSvsCAMA/10-26-26/Photos_New"

echo.
echo ============================================================
echo  DONE - Close this window when finished
echo ============================================================
pause
