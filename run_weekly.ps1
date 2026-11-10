# Weekly system check (run_weekly.bat runs this file):
#   [1/4] collect event log CSVs     EventLogCheck.ps1 / EventLogCheck_Local.ps1, paths from config.json
#   [2/4] export monitor data        Kettle job (Kitchen, run from the Kettle folder), log -> <log_dir>\yyyyMMdd_kettle.log
#   [3/4] monitor report             monitor_check
#   [4/4] fill check sheet and tracking sheet   fill_check
# All settings come from config.json in this folder:
#   weekly.collector / weekly.kitchen / weekly.kettle_job
#   paths.csv_dir -> -OutDir, log_dir -> -LogDir, eventlog.server_list / exclude_file / days / start_time / end_time
# Failure handling:
#   [1/4] fails (exit 1)          -> stop, nothing else runs
#   [2/4] Kettle fails            -> fill_check still runs with --clear-monitor (monitor columns cleared)
#   [3/4] monitor_check fails     -> warning only: fill_check reads the Kettle CSVs itself, the report is just for reading
# Exit code: 1 = config or event log collection failed, otherwise fill_check's exit code (3 = Excel file open)
# Keep this file pure ASCII (Windows PowerShell 5.1 reads BOM-less files as ANSI)
param(
    [string]$Base = $PSScriptRoot
)

$config = Join-Path $Base 'config.json'
try {
    # config.json is UTF-8 (Chinese folder names and \\server\share paths are OK)
    $cfg = Get-Content -LiteralPath $config -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json
}
catch {
    Write-Host "[ERROR] Cannot read $config : $($_.Exception.Message)"
    exit 1
}
Write-Host "Config: $config"

# D:/SystemCheck/logs -> D:\SystemCheck\logs, relative -> under Base; empty -> $null
function Resolve-ConfigPath($p) {
    if ([string]::IsNullOrWhiteSpace([string]$p)) { return $null }
    $p = ([string]$p).Replace('/', '\')
    if (-not [IO.Path]::IsPathRooted($p)) { $p = Join-Path $Base $p }
    return $p.TrimEnd('\')
}

# .NET Framework has no ISOWeek and GetWeekOfYear differs around New Year, so compute it:
# the ISO week/year are those of the Thursday in the same Monday-Sunday week.
function Get-IsoWeek([datetime]$d) {
    $thu  = $d.Date.AddDays(3 - (([int]$d.DayOfWeek + 6) % 7))
    $week = [math]::Floor(($thu.DayOfYear - 1) / 7) + 1
    return @{ yyyy = $thu.Year.ToString(); yy = '{0:D2}' -f ($thu.Year % 100); ww = '{0:D2}' -f [int]$week }
}

# Command for a Python tool: <exe> if present, otherwise python <name>.py; $null when neither works.
# (The tool itself is run at script level: output of a command run inside a function would become
#  the function's return value.)
function Get-ToolCommand([string]$name, [string[]]$exes) {
    $exe = $exes | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
    if ($exe) { return , @($exe) }
    if (Get-Command python -ErrorAction SilentlyContinue) { return , @('python', (Join-Path $Base "$name.py")) }
    Write-Host "[ERROR] $name.exe not found ($($exes -join ', ')) and Python is not installed"
    return $null
}

$weekly    = $cfg.weekly
$collector = if ($weekly.collector) { [string]$weekly.collector } else { 'EventLogCheck.ps1' }
$kitchen   = Resolve-ConfigPath $weekly.kitchen
$kettleJob = Resolve-ConfigPath $weekly.kettle_job
$failed    = @()
$reportFailed = $false

# ---------- [1/4] event logs ----------
Write-Host ""
Write-Host "[1/4] Collecting event logs with $collector ..."
$script = Join-Path $Base $collector
if (-not (Test-Path -LiteralPath $script -PathType Leaf)) {
    Write-Host "[ERROR] Collector not found: $script - check weekly.collector in config.json"
    exit 1
}
$el = $cfg.eventlog
$values = [ordered]@{
    OutDir      = Resolve-ConfigPath $cfg.paths.csv_dir
    LogDir      = Resolve-ConfigPath $cfg.log_dir
    ServerList  = Resolve-ConfigPath $el.server_list
    ExcludeFile = Resolve-ConfigPath $el.exclude_file
    Days        = [string]$el.days
    StartTime   = [string]$el.start_time
    EndTime     = [string]$el.end_time
}
# Only parameters the collector has (EventLogCheck_Local.ps1 has no -ServerList); empty settings are not passed
$accepted = (Get-Command -Name $script).Parameters.Keys
$params = [ordered]@{}
foreach ($k in $values.Keys) {
    if (-not [string]::IsNullOrWhiteSpace($values[$k]) -and $accepted -contains $k) { $params[$k] = $values[$k] }
}
foreach ($k in $params.Keys) { Write-Host ("  -{0,-11} {1}" -f $k, $params[$k]) }
& $script @params
$code = $LASTEXITCODE
if ($code -eq 2) {
    # No errors/warnings left after exclusion: no CSV. fill_check still creates this week's check sheet.
    Write-Host "No error/warning events after exclusion - no CSV written this time."
}
elseif ($code -ne 0) {
    Write-Host ""
    Write-Host "[ERROR] Event log collection failed. Stopped before Kettle and filling the Excel files."
    Write-Host "        See the messages above and the log folder - log_dir in config.json"
    exit 1
}

# ---------- [2/4] Kettle ----------
Write-Host ""
Write-Host "[2/4] Exporting monitor data with Kettle ..."
$kettleOk = $false
if (-not $kitchen -or -not (Test-Path -LiteralPath $kitchen -PathType Leaf)) {
    Write-Host "[ERROR] Kettle not found: $kitchen - set weekly.kitchen in config.json"
}
elseif (-not $kettleJob -or -not (Test-Path -LiteralPath $kettleJob -PathType Leaf)) {
    Write-Host "[ERROR] Kettle job not found: $kettleJob - set weekly.kettle_job in config.json"
}
else {
    # Kettle log: <log_dir>\yyyyMMdd_kettle.log ({week} etc. = ISO week of today)
    $logDir = [string]$cfg.log_dir
    if ([string]::IsNullOrWhiteSpace($logDir)) { $logDir = 'logs' }
    $w = Get-IsoWeek (Get-Date)
    $logDir = $logDir.Replace('{week}', "W$($w.yy)$($w.ww)").Replace('{yyyy}', $w.yyyy).Replace('{yy}', $w.yy).Replace('{ww}', $w.ww)
    $logDir = [IO.Path]::GetFullPath((Resolve-ConfigPath $logDir))
    try {
        New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop | Out-Null
        $kettleLog = Join-Path $logDir ((Get-Date -Format 'yyyyMMdd') + '_kettle.log')
        Write-Host "Kettle log: $kettleLog"
        # Kitchen.bat / Spoon.bat use relative paths (launcher\launcher.jar, libswt\win64):
        # run them from the Kettle folder; job and log are absolute paths
        Push-Location -LiteralPath (Split-Path -Parent $kitchen)
        try {
            & $kitchen "/file:$kettleJob" '/level:Basic' "/logfile:$kettleLog"
            $kitchenCode = $LASTEXITCODE
        }
        finally {
            Pop-Location
        }
        if ($kitchenCode -eq 0) { $kettleOk = $true }
        else { Write-Host "[ERROR] Kitchen exit code $kitchenCode - see $kettleLog" }
    }
    catch {
        Write-Host "[ERROR] Cannot create log folder $logDir : $($_.Exception.Message)"
    }
}
if (-not $kettleOk) {
    Write-Host "[ERROR] Kettle failed - monitor_check skipped, monitor columns in the check sheet will be cleared"
    $failed += 'Kettle'
}

# ---------- [3/4] monitor_check ----------
if ($kettleOk) {
    Write-Host ""
    Write-Host "[3/4] Checking monitor data with monitor_check ..."
    $tool = Get-ToolCommand 'monitor_check' @((Join-Path $Base 'monitor_check.exe'))
    $code = 1
    if ($tool) {
        & $tool[0] @($tool | Select-Object -Skip 1)
        $code = $LASTEXITCODE
    }
    if ($code -ne 0) {
        # fill_check analyses the Kettle CSVs itself, so the check sheet does not depend on this report
        Write-Host "[WARN] monitor_check failed (exit code $code) - no monitor report; the check sheet is still filled from the Kettle CSVs"
        $reportFailed = $true
    }
}

# ---------- [4/4] fill_check ----------
Write-Host ""
Write-Host "[4/4] Filling check sheet and tracking sheet..."
# When Kettle failed, --clear-monitor empties the monitor columns (CPU, Memory, ...)
$fillArgs = @()
if ($failed.Count -gt 0) { $fillArgs += '--clear-monitor' }
# fill_check.exe (onefile) next to this script; the old onedir layout fill_check\fill_check.exe still works
$tool = Get-ToolCommand 'fill_check' @((Join-Path $Base 'fill_check.exe'), (Join-Path $Base 'fill_check\fill_check.exe'))
$code = 1
if ($tool) {
    & $tool[0] @(@($tool | Select-Object -Skip 1) + $fillArgs)
    $code = $LASTEXITCODE
}
if ($code -eq 3) {
    Write-Host "[ERROR] An Excel file was open by someone - not updated. Run again later, it is safe to re-run."
}
elseif ($code -ne 0) {
    Write-Host "[ERROR] fill_check did not finish. See its log."
}
else {
    Write-Host "Done."
}
if ($failed.Count -gt 0) {
    Write-Host ""
    Write-Host "[WARN] Failed steps: $($failed -join ' ') - monitor columns were cleared, see the messages above"
}
if ($reportFailed) {
    Write-Host "[WARN] monitor_check failed (report only) - see its log: <log_dir>\yyyyMMdd_monitor_check.log"
}
exit $code
