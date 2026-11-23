@echo off
cd /d "%~dp0"

echo ============================================================
echo  Zillow Photo Downloader - Value Mismatches
echo ============================================================
py "ZillowPhotos/download_zillow_photos.py" "MLSvsCAMA/11-23-26/value_mismatches_2026-11-23.xlsx" "MLSvsCAMA/11-23-26/Photos_New"

echo.
echo ============================================================
echo  Zillow Photo Downloader - Perfect Matches
echo ============================================================
py "ZillowPhotos/download_zillow_photos.py" "MLSvsCAMA/11-23-26/perfect_matches_2026-11-23.xlsx" "MLSvsCAMA/11-23-26/Photos_New"

echo.
echo ============================================================
echo  DONE - Close this window when finished
echo ============================================================
pause
