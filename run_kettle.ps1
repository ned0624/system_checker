# Run the Kettle monitor job with Kitchen and write its log to log_dir of config.json:
#   <log_dir>\yyyyMMdd_kettle.log
# Called by run_weekly.bat step [2/4]:
#   powershell -NoProfile -ExecutionPolicy Bypass -File run_kettle.ps1 -Base D:\SystemCheck -Kitchen ...\Kitchen.bat -Job ...\monitor.kjb
# - log_dir may contain {week} (W2645), {yyyy} (2026), {yy} (26), {ww} (45) = ISO week of today;
#   a relative log_dir is relative to Base; empty or missing log_dir = <Base>\logs
# - Exit code: Kitchen's exit code (0 = success), 1 = config / Kitchen / job not found
# - Keep this file pure ASCII (Windows PowerShell 5.1 reads BOM-less files as ANSI)
param(
    [Parameter(Mandatory = $true)][string]$Base,
    [Parameter(Mandatory = $true)][string]$Kitchen,
    [Parameter(Mandatory = $true)][string]$Job
)

# .NET Framework has no ISOWeek and GetWeekOfYear differs around New Year, so compute it:
# the ISO week/year are those of the Thursday in the same Monday-Sunday week.
function Get-IsoWeek([datetime]$d) {
    $thu  = $d.Date.AddDays(3 - (([int]$d.DayOfWeek + 6) % 7))
    $week = [math]::Floor(($thu.DayOfYear - 1) / 7) + 1
    return @{ yyyy = $thu.Year.ToString(); yy = '{0:D2}' -f ($thu.Year % 100); ww = '{0:D2}' -f [int]$week }
}

$config = Join-Path $Base 'config.json'
try {
    # config.json is UTF-8 (Chinese folder names are OK)
    $cfg = Get-Content -LiteralPath $config -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json
}
catch {
    Write-Host "[ERROR] Cannot read $config : $($_.Exception.Message)"
    exit 1
}

$logDir = [string]$cfg.log_dir
if ([string]::IsNullOrWhiteSpace($logDir)) { $logDir = 'logs' }
$w = Get-IsoWeek (Get-Date)
$logDir = $logDir.Replace('{week}', "W$($w.yy)$($w.ww)").Replace('{yyyy}', $w.yyyy).Replace('{yy}', $w.yy).Replace('{ww}', $w.ww)
if (-not [IO.Path]::IsPathRooted($logDir)) { $logDir = Join-Path $Base $logDir }
$logDir = [IO.Path]::GetFullPath($logDir)  # D:/SystemCheck/logs -> D:\SystemCheck\logs
try {
    New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop | Out-Null
}
catch {
    Write-Host "[ERROR] Cannot create log folder $logDir : $($_.Exception.Message)"
    exit 1
}
$log = Join-Path $logDir ((Get-Date -Format 'yyyyMMdd') + '_kettle.log')

if (-not (Test-Path -LiteralPath $Kitchen -PathType Leaf)) {
    Write-Host "[ERROR] Kettle not found: $Kitchen - set KITCHEN in run_weekly.bat"
    exit 1
}
if (-not (Test-Path -LiteralPath $Job -PathType Leaf)) {
    Write-Host "[ERROR] Kettle job not found: $Job - set KETTLE_JOB in run_weekly.bat"
    exit 1
}

Write-Host "Kettle log: $log"
& $Kitchen "/file:$Job" '/level:Basic' "/logfile:$log"
$code = $LASTEXITCODE
if ($code -ne 0) { Write-Host "[ERROR] Kitchen exit code $code - see $log" }
exit $code
