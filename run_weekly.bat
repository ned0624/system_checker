@echo off
REM Weekly system check - double-click to run, or add to Task Scheduler.
REM run_weekly.ps1 (same folder) does all the steps:
REM   [1/4] collect event log CSVs  [2/4] Kettle  [3/4] monitor_check  [4/4] fill_check
REM All settings are in config.json (same folder): paths, log_dir, eventlog, monitor, weekly
REM   weekly.collector  = EventLogCheck.ps1 (server list, remote) or EventLogCheck_Local.ps1 (this computer)
REM   weekly.kitchen    = Kettle Kitchen.bat
REM   weekly.kettle_job = Kettle monitor job (.kjb)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_weekly.ps1"
echo.
pause
