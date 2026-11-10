# Weekly system check (run_weekly.bat runs this file):
#   [1/4] collect event log CSVs     EventLogCheck.ps1 / EventLogCheck_Local.ps1, paths from config.json
#   [2/4] export monitor data        Kettle job (Kitchen, run from the Kettle folder)
#   [3/4] monitor report             monitor_check
#   [4/4] fill check sheet and tracking sheet   fill_check
# All settings come from config.json in this folder:
#   weekly.collector / weekly.kitchen / weekly.kettle_job
#   paths.csv_dir -> -OutDir, log_dir -> -LogDir, eventlog.server_list / exclude_file / days / start_time / end_time
# Failure handling:
#   [1/4] fails (exit 1)          -> stop, nothing else runs
#   [2/4] Kettle fails            -> fill_check still runs with --clear-monitor (monitor columns cleared)
#   [3/4] monitor_check fails     -> warning only: fill_check reads the Kettle CSVs itself, the report is just for reading
# Log: everything goes to one file <log_dir>\yyyyMMdd.log - these messages, EventLogCheck.ps1, the Kettle log,
#      monitor_check and fill_check (all append to the same file)
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
    Write-Log "[ERROR] $name.exe not found ($($exes -join ', ')) and Python is not installed"
    return $null
}

# One log file for the whole run: <log_dir>\yyyyMMdd.log ({week} etc. = ISO week of today)
$logDir = [string]$cfg.log_dir
if ([string]::IsNullOrWhiteSpace($logDir)) { $logDir = 'logs' }
$w = Get-IsoWeek (Get-Date)
$logDir = $logDir.Replace('{week}', "W$($w.yy)$($w.ww)").Replace('{yyyy}', $w.yyyy).Replace('{yy}', $w.yy).Replace('{ww}', $w.ww)
$logDir = [IO.Path]::GetFullPath((Resolve-ConfigPath $logDir))
$logFile = Join-Path $logDir ((Get-Date -Format 'yyyyMMdd') + '.log')
try {
    New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop | Out-Null
}
catch {
    Write-Host "[WARN] Cannot create log folder $logDir : $($_.Exception.Message) - messages are shown on screen only"
    $logFile = $null
}
$utf8Bom = New-Object Text.UTF8Encoding($true)   # BOM only when the file is new

# Show a message and append it to the log file
function Write-Log([string]$msg) {
    Write-Host $msg
    if ($script:logFile) {
        try { [IO.File]::AppendAllText($script:logFile, $msg + "`r`n", $script:utf8Bom) } catch { }
    }
}

Write-Log ""
Write-Log "===== run_weekly start $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ====="
Write-Log "Config: $config"
if ($logFile) { Write-Log "Log   : $logFile" }

$weekly    = $cfg.weekly
$collector = if ($weekly.collector) { [string]$weekly.collector } else { 'EventLogCheck.ps1' }
$kitchen   = Resolve-ConfigPath $weekly.kitchen
$kettleJob = Resolve-ConfigPath $weekly.kettle_job
$failed    = @()
$reportFailed = $false

# ---------- [1/4] event logs ----------
Write-Log ""
Write-Log "[1/4] Collecting event logs with $collector ..."
$script = Join-Path $Base $collector
if (-not (Test-Path -LiteralPath $script -PathType Leaf)) {
    Write-Log "[ERROR] Collector not found: $script - check weekly.collector in config.json"
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
foreach ($k in $params.Keys) { Write-Log ("  -{0,-11} {1}" -f $k, $params[$k]) }
& $script @params
$code = $LASTEXITCODE
if ($code -eq 2) {
    # No errors/warnings left after exclusion: no CSV. fill_check still creates this week's check sheet.
    Write-Log "No error/warning events after exclusion - no CSV written this time."
}
elseif ($code -ne 0) {
    Write-Log ""
    Write-Log "[ERROR] Event log collection failed. Stopped before Kettle and filling the Excel files."
    Write-Log "        See the messages above and the log folder - log_dir in config.json"
    exit 1
}

# ---------- [2/4] Kettle ----------
Write-Log ""
Write-Log "[2/4] Exporting monitor data with Kettle ..."
$kettleOk = $false
if (-not $kitchen -or -not (Test-Path -LiteralPath $kitchen -PathType Leaf)) {
    Write-Log "[ERROR] Kettle not found: $kitchen - set weekly.kitchen in config.json"
}
elseif (-not $kettleJob -or -not (Test-Path -LiteralPath $kettleJob -PathType Leaf)) {
    Write-Log "[ERROR] Kettle job not found: $kettleJob - set weekly.kettle_job in config.json"
}
else {
    # Kitchen writes its log to a temp file; it is appended to the day log afterwards
    # (Kettle writes the file in the Windows ANSI code page, the day log is UTF-8)
    $kettleTmp = Join-Path $env:TEMP ('kettle_' + [guid]::NewGuid().ToString('N') + '.log')
    # Kitchen.bat / Spoon.bat use relative paths (launcher\launcher.jar, libswt\win64):
    # run them from the Kettle folder; job and log are absolute paths
    Push-Location -LiteralPath (Split-Path -Parent $kitchen)
    try {
        & $kitchen "/file:$kettleJob" '/level:Basic' "/logfile:$kettleTmp"
        $kitchenCode = $LASTEXITCODE
    }
    finally {
        Pop-Location
    }
    if (Test-Path -LiteralPath $kettleTmp) {
        $kettleText = [IO.File]::ReadAllText($kettleTmp, [Text.Encoding]::Default)
        if ($logFile) {
            try {
                [IO.File]::AppendAllText($logFile, "----- Kettle log -----`r`n" + $kettleText.TrimEnd() + "`r`n----- end of Kettle log -----`r`n", $utf8Bom)
            } catch { }
        }
        Remove-Item -LiteralPath $kettleTmp -Force -ErrorAction SilentlyContinue
    }
    if ($kitchenCode -eq 0) { $kettleOk = $true }
    else { Write-Log "[ERROR] Kitchen exit code $kitchenCode - see the Kettle log in $logFile" }
}
if (-not $kettleOk) {
    Write-Log "[ERROR] Kettle failed - monitor_check skipped, monitor columns in the check sheet will be cleared"
    $failed += 'Kettle'
}

# ---------- [3/4] monitor_check ----------
if ($kettleOk) {
    Write-Log ""
    Write-Log "[3/4] Checking monitor data with monitor_check ..."
    $tool = Get-ToolCommand 'monitor_check' @((Join-Path $Base 'monitor_check.exe'))
    $code = 1
    if ($tool) {
        & $tool[0] @($tool | Select-Object -Skip 1)
        $code = $LASTEXITCODE
    }
    if ($code -ne 0) {
        # fill_check analyses the Kettle CSVs itself, so the check sheet does not depend on this report
        Write-Log "[WARN] monitor_check failed (exit code $code) - no monitor report; the check sheet is still filled from the Kettle CSVs"
        $reportFailed = $true
    }
}

# ---------- [4/4] fill_check ----------
Write-Log ""
Write-Log "[4/4] Filling check sheet and tracking sheet..."
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
    Write-Log "[ERROR] An Excel file was open by someone - not updated. Run again later, it is safe to re-run."
}
elseif ($code -ne 0) {
    Write-Log "[ERROR] fill_check did not finish. See its log."
}
else {
    Write-Log "Done."
}
if ($failed.Count -gt 0) {
    Write-Log ""
    Write-Log "[WARN] Failed steps: $($failed -join ' ') - monitor columns were cleared, see the messages above"
}
if ($reportFailed) {
    Write-Log "[WARN] monitor_check failed (report only) - see $logFile"
}
exit $code
