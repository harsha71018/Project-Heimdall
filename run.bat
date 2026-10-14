@echo off
title "Automated Job Hunter and Excel Dashboard"
color 0B

echo =====================================================================
echo       AUTOMATED JOB DISCOVERY AND APPLICATION TRACKER ENGINE
echo =====================================================================
echo.

cd /d "%~dp0"

echo [1/2] Verifying Python Environment...
py -c "import openpyxl, pandas, jobspy, yaml, dotenv, requests" >nul 2>&1
if errorlevel 1 (
    echo Installing or updating required packages...
    py -m pip install -r requirements.txt --quiet --disable-pip-version-check
    if errorlevel 1 (
        echo [!] Error: Failed to install required Python packages.
        pause
        exit /b 1
    )
) else (
    echo Python environment verified.
)

echo.
echo =====================================================================
echo Select Scraping Intensity for Today:
echo   [1] Balanced   - Recommended: Full Markdown on top listings
echo   [2] Frugal     - Quick search scan
echo   [3] Exhaustive - Deep sweep
echo =====================================================================
choice /c 123 /t 5 /d 1 /m "Auto-selecting [1] in 5 seconds"
if "%ERRORLEVEL%"=="3" (
    set MODE=exhaustive
) else if "%ERRORLEVEL%"=="2" (
    set MODE=frugal
) else (
    set MODE=balanced
)

echo.
echo [2/2] Launching Job Hunter Engine in [%MODE%] mode...
echo.
py main.py --mode %MODE%

echo.
echo =====================================================================
echo Execution Finished. Press any key to exit...
pause >nul
