@echo off
REM Weekly system check: 1. collect event log CSVs  2. fill check sheet and tracking sheet
REM Double-click to run, or add to Task Scheduler
set BASE=D:\SystemCheck

REM Which collector to use:
REM   EventLogCheck.ps1        = collect from every server in the server list (remote)
REM   EventLogCheck_Local.ps1  = collect from this computer only
REM Folders (CSV output, log, server list, exclude list) are set in the SETTINGS block of the ps1
set COLLECTOR=EventLogCheck.ps1

echo [1/2] Collecting event logs with %COLLECTOR% ...
powershell -NoProfile -ExecutionPolicy Bypass -File "%BASE%\%COLLECTOR%"
REM Exit code 2 = no errors/warnings left after exclusion, no CSV written: nothing to fill in
REM (check 2 first: "if errorlevel N" means "exit code >= N")
if errorlevel 2 (
    echo.
    echo No error/warning events after exclusion - no CSV written, nothing to fill in.
    goto END
)
REM Exit code 1 = error (folder not writable, or every server failed)
if errorlevel 1 (
    echo.
    echo [ERROR] Event log collection failed. Stopped before filling the Excel files.
    echo         See the messages above and the log folder set in %COLLECTOR%
    goto END
)

echo [2/2] Filling check sheet and tracking sheet...
REM fill_check reads CSVs from paths.csv_dir in config.json - it must be the same folder as OutDir in the ps1
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
    echo [ERROR] fill_check did not finish. See its log.
) else (
    echo Done.
)

:END
echo.
pause
