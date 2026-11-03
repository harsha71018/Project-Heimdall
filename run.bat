@echo off
title "Project Heimdall — Automated Job Hunter & Excel Dashboard"
color 0B

echo =====================================================================
echo       PROJECT HEIMDALL — AUTOMATED JOB DISCOVERY ^& TRACKER ENGINE
echo =====================================================================
echo.

cd /d "%~dp0"

if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
    set "PATH=%LOCALAPPDATA%\Programs\Python\Python313;%LOCALAPPDATA%\Programs\Python\Python313\Scripts;%PATH%"
)

set PYTHON_CMD=py
py -V >nul 2>&1
if errorlevel 1 (
    set PYTHON_CMD=python
)

echo [1/2] Verifying Python Environment...
%PYTHON_CMD% -c "import openpyxl, pandas, jobspy, yaml, dotenv, requests" >nul 2>&1
if errorlevel 1 (
    echo Installing or updating required packages...
    %PYTHON_CMD% -m pip install -r requirements.txt --quiet --disable-pip-version-check
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
if not "%~1"=="" (
    set "MODE=%~1"
    goto launch
)
if not "%MODE%"=="" goto launch
if "%NONINTERACTIVE%"=="1" (
    set "MODE=balanced"
    goto launch
)

echo Select Scraping Intensity for Today:
echo   [1] Balanced   (~24 credits) - Recommended: Full Markdown on top 4 listings
echo   [2] Frugal     (  8 credits) - Quick search scan
echo   [3] Exhaustive (~48 credits) - Deep Monday sweep
echo =====================================================================
choice /c 123 /t 5 /d 1 /m "Auto-selecting [1] in 5 seconds"
if "%ERRORLEVEL%"=="3" (
    set "MODE=exhaustive"
) else if "%ERRORLEVEL%"=="2" (
    set "MODE=frugal"
) else (
    set "MODE=balanced"
)

:launch
echo.
echo [2/2] Launching Job Hunter Engine in [%MODE%] mode...
echo.
%PYTHON_CMD% main.py --mode %MODE%

echo.
echo =====================================================================
echo Execution Finished. Excel tracker updated: Tracker_Latest.xlsx
if not "%NONINTERACTIVE%"=="1" pause >nul
