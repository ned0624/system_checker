# Run the event log collector (EventLogCheck.ps1 or EventLogCheck_Local.ps1) with the paths from config.json.
# Called by run_weekly.bat step [1/4]:
#   powershell -NoProfile -ExecutionPolicy Bypass -File run_collector.ps1 -Base D:\SystemCheck -Collector EventLogCheck.ps1
# config.json -> collector parameters:
#   paths.csv_dir          -> -OutDir       (the same folder fill_check reads)
#   log_dir                -> -LogDir
#   eventlog.server_list   -> -ServerList   (EventLogCheck.ps1 only)
#   eventlog.exclude_file  -> -ExcludeFile
#   eventlog.days / start_time / end_time -> -Days / -StartTime / -EndTime
# - Settings that are missing or empty are not passed: the collector then uses its own SETTINGS block
# - Relative paths are relative to Base; {week} etc. are replaced by the collector (ISO week of the period end)
# - Exit code: the collector's exit code (0 = CSV written, 1 = error, 2 = nothing to report)
# - Keep this file pure ASCII (Windows PowerShell 5.1 reads BOM-less files as ANSI)
param(
    [Parameter(Mandatory = $true)][string]$Base,
    [Parameter(Mandatory = $true)][string]$Collector
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

$script = Join-Path $Base $Collector
if (-not (Test-Path -LiteralPath $script -PathType Leaf)) {
    Write-Host "[ERROR] Collector not found: $script - check COLLECTOR in run_weekly.bat"
    exit 1
}

# D:/SystemCheck/logs -> D:\SystemCheck\logs, relative -> under Base; empty -> $null (not passed)
function Resolve-ConfigPath($p) {
    if ([string]::IsNullOrWhiteSpace([string]$p)) { return $null }
    $p = ([string]$p).Replace('/', '\')
    if (-not [IO.Path]::IsPathRooted($p)) { $p = Join-Path $Base $p }
    return $p.TrimEnd('\')
}

$params = [ordered]@{}
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
# Only parameters the collector has (EventLogCheck_Local.ps1 has no -ServerList)
$accepted = (Get-Command -Name $script).Parameters.Keys
foreach ($k in $values.Keys) {
    if (-not [string]::IsNullOrWhiteSpace($values[$k]) -and $accepted -contains $k) { $params[$k] = $values[$k] }
}

Write-Host "Config: $config"
foreach ($k in $params.Keys) { Write-Host ("  -{0,-11} {1}" -f $k, $params[$k]) }
& $script @params
exit $LASTEXITCODE
