@echo off
cd /d "%~dp0"

if exist "scheduled_run.log" (
    for %%A in ("scheduled_run.log") do (
        if %%~zA GTR 5242880 move /y "scheduled_run.log" "scheduled_run.log.old" >nul 2>&1
    )
)

echo. >> scheduled_run.log
echo ===================================================================== >> scheduled_run.log
echo [%DATE% %TIME%] Starting Scheduled Frugal Job Sweep >> scheduled_run.log
echo ===================================================================== >> scheduled_run.log
py -u main.py --mode frugal >> scheduled_run.log 2>&1
