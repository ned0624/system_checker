@echo off
REM Weekly system check:
REM   1. collect event log CSVs          (EventLogCheck.ps1)
REM   2. export monitor data             (Kettle job, reads D:\SystemCheck\config.json)
REM   3. monitor report                  (monitor_check)
REM   4. fill check sheet and tracking sheet (fill_check)
REM Double-click to run, or add to Task Scheduler
set BASE=D:\SystemCheck

REM Which collector to use:
REM   EventLogCheck.ps1        = collect from every server in the server list (remote)
REM   EventLogCheck_Local.ps1  = collect from this computer only
REM Folders (CSV output, log, server list, exclude list) are set in the SETTINGS block of the ps1
set COLLECTOR=EventLogCheck.ps1

REM Kettle (Pentaho Data Integration): Kitchen.bat and the monitor job (read_config.ktr + monitor.ktr)
set KITCHEN=D:\data-integration\Kitchen.bat
set KETTLE_JOB=%BASE%\kettle\monitor.kjb

set MONITOR_FAILED=
set FAILED_STEPS=

echo [1/4] Collecting event logs with %COLLECTOR% ...
powershell -NoProfile -ExecutionPolicy Bypass -File "%BASE%\%COLLECTOR%"
REM Exit code 2 = no errors/warnings left after exclusion, no CSV written.
REM The next steps still run: fill_check only uses this week's CSVs and creates this week's check sheet (all O).
REM (check 2 first: "if errorlevel N" means "exit code >= N")
if errorlevel 2 (
    echo No error/warning events after exclusion - no CSV written this time.
    goto KETTLE
)
REM Exit code 1 = error (folder not writable, or every server failed)
if errorlevel 1 (
    echo.
    echo [ERROR] Event log collection failed. Stopped before Kettle and filling the Excel files.
    echo         See the messages above and the log folder set in %COLLECTOR%
    goto END
)

:KETTLE
echo.
echo [2/4] Exporting monitor data with Kettle ...
if not exist "%KITCHEN%" (
    echo [ERROR] Kettle not found: %KITCHEN% - set KITCHEN in run_weekly.bat
    goto KETTLE_FAILED
)
if not exist "%KETTLE_JOB%" (
    echo [ERROR] Kettle job not found: %KETTLE_JOB% - set KETTLE_JOB in run_weekly.bat
    goto KETTLE_FAILED
)
call "%KITCHEN%" /file:"%KETTLE_JOB%" /level:Basic
if errorlevel 1 goto KETTLE_FAILED
goto MONITOR

:KETTLE_FAILED
echo [ERROR] Kettle failed - monitor_check skipped, monitor columns in the check sheet will be cleared
set MONITOR_FAILED=1
set FAILED_STEPS=%FAILED_STEPS% Kettle
goto FILL

:MONITOR
echo.
echo [3/4] Checking monitor data with monitor_check ...
REM Use monitor_check.exe if present, otherwise run monitor_check.py with Python
if exist "%BASE%\monitor_check.exe" (
    "%BASE%\monitor_check.exe"
) else (
    python "%BASE%\monitor_check.py"
)
if errorlevel 1 (
    echo [ERROR] monitor_check failed - monitor columns in the check sheet will be cleared
    set MONITOR_FAILED=1
    set FAILED_STEPS=%FAILED_STEPS% monitor_check
)

:FILL
echo.
echo [4/4] Filling check sheet and tracking sheet...
REM fill_check reads this week's CSVs from paths.csv_dir in config.json - it must be the same folder as OutDir in the ps1
REM When Kettle or monitor_check failed, --clear-monitor empties the monitor columns (CPU, Memory, ...)
set FILL_ARGS=
if defined MONITOR_FAILED set FILL_ARGS=--clear-monitor
REM Use fill_check.exe if present, otherwise run fill_check.py with Python
if exist "%BASE%\fill_check\fill_check.exe" (
    "%BASE%\fill_check\fill_check.exe" %FILL_ARGS%
) else (
    where python >nul 2>nul
    if errorlevel 1 (
        echo [ERROR] %BASE%\fill_check\fill_check.exe not found and Python is not installed
        goto END
    )
    python "%BASE%\fill_check.py" %FILL_ARGS%
)
REM fill_check exit code 3 = an Excel file stayed open by someone longer than wait_if_open allows
if errorlevel 3 (
    echo [ERROR] An Excel file was open by someone - not updated. Run again later, it is safe to re-run.
) else if errorlevel 1 (
    echo [ERROR] fill_check did not finish. See its log.
) else (
    echo Done.
)
if defined FAILED_STEPS (
    echo.
    echo [WARN] Failed steps:%FAILED_STEPS% - monitor columns were cleared, see the messages above
)

:END
echo.
pause
