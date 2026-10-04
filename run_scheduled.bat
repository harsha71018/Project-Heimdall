@echo off
cd /d "%~dp0"
py main.py --mode frugal >> scheduled_run.log 2>&1
