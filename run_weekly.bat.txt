@echo off
REM Weekly system check: 1. collect event log CSVs  2. fill check sheet and tracking sheet
REM Double-click to run, or add to Task Scheduler
set BASE=D:\SystemCheck

for /f %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set TODAY=%%d

echo [1/2] Collecting event logs...
powershell -NoProfile -ExecutionPolicy Bypass -File "%BASE%\EventLogCheck.ps1"

REM Stop if no CSV for today, so fill_check does not re-run last week's CSVs
REM and overwrite last week's tracking sheet
if not exist "%BASE%\EventLog\%TODAY%_*.csv" (
    echo [ERROR] No CSV created for today %TODAY%. Stopped. Check server connections and serverlist.txt
    goto END
)

echo [2/2] Filling check sheet and tracking sheet...
REM Use fill_check.exe if present, otherwise run fill_check.py with Python
if exist "%BASE%\fill_check\fill_check.exe" (
    "%BASE%\fill_check\fill_check.exe"
) else (
    where python >nul 2>nul
    if errorlevel 1 (
        echo [ERROR] %BASE%\fill_check\fill_check.exe not found and Python is not installed
        goto END
    )
    python "%BASE%\fill_check.py"
)
if errorlevel 1 (
    echo [ERROR] fill_check did not finish. See the log in %BASE%\logs
) else (
    echo Done. See the log in %BASE%\logs
)

:END
echo.
pause
