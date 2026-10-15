@echo off
title "IT Support & ServiceNow Job Hunter"
color 0A

echo =====================================================================
echo       IT SUPPORT & SERVICENOW JOB DISCOVERY ENGINE
echo       Target: Hyderabad & Bengaluru | 0-2 YOE
echo =====================================================================
echo.

cd /d "%~dp0"

echo [1/2] Verifying Python Environment...
py -c "import openpyxl, pandas, jobspy, yaml, dotenv, requests" >nul 2>&1
if errorlevel 1 (
    echo Installing or updating required packages...
    py -m pip install -r requirements.txt --quiet --disable-pip-version-check
    if errorlevel 1 (
        py -m pip install numpy pandas openpyxl pyyaml requests beautifulsoup4 tls-client markdownify pydantic python-dotenv --quiet --disable-pip-version-check
        py -m pip install python-jobspy==1.1.82 --no-deps --quiet --disable-pip-version-check
    )
) else (
    echo Python environment verified.
)

echo.
echo =====================================================================
echo Select Scraping Intensity for IT Support:
echo   [1] Frugal     - Fast search scan (Recommended: ~32 credits total)
echo   [2] Balanced   - Full Markdown on top listings (~80 credits)
echo   [3] Exhaustive - Deep sweep
echo =====================================================================
choice /c 123 /t 5 /d 1 /m "Auto-selecting [1] in 5 seconds"
if "%ERRORLEVEL%"=="3" (
    set MODE=exhaustive
) else if "%ERRORLEVEL%"=="2" (
    set MODE=balanced
) else (
    set MODE=frugal
)

echo.
echo [2/2] Launching IT Support Job Hunter in [%MODE%] mode...
echo.
py main_it_support.py --config config_it_support.yaml --mode %MODE%

echo.
echo =====================================================================
echo Execution Finished! Excel tracker updated: Tracker_Latest_IT_Support.xlsx
echo Press any key to exit...
pause >nul
