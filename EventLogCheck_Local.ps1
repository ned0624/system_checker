# Collect event log errors/warnings of the last N days on THIS computer, one CSV:
#   <OutDir>\yyyyMMdd_<computer name>.csv
# Usage (run on each server): powershell -ExecutionPolicy Bypass -File EventLogCheck_Local.ps1 [-OutDir ...] [-LogDir ...] [-Days 7]
# - OutDir can be a share (e.g. \\fileserver\SystemCheck\EventLog) to collect all servers' CSVs for fill_check
# - Only Critical/Error/Warning (Level 1-3), no Information; header-only CSV when there are none
# - Keep this file pure ASCII (see $levelNames below)
param(
    [string]$OutDir = 'D:\SystemCheck\EventLog',
    [string]$LogDir = 'D:\SystemCheck\logs',
    [int]$Days      = 7
)

$ErrorActionPreference = 'Continue'
$start = (Get-Date).AddDays(-$Days)
$date  = Get-Date -Format 'yyyyMMdd'
New-Item -ItemType Directory -Force -Path $OutDir, $LogDir | Out-Null
$logFile = Join-Path $LogDir "$date.log"

# Level name from the level number, not LevelDisplayName (empty for some providers).
# Chinese names are built from char codes so this file stays pure ASCII:
# Windows PowerShell 5.1 reads a .ps1 without BOM as ANSI, which would garble Chinese literals.
#   1 = Critical (0x91CD 0x5927), 2 = Error (0x932F 0x8AA4), 3 = Warning (0x8B66 0x544A)
$levelNames = @{
    1 = [string][char]0x91CD + [char]0x5927
    2 = [string][char]0x932F + [char]0x8AA4
    3 = [string][char]0x8B66 + [char]0x544A
}
$csvHeader  = '"LevelDisplayName","Id","LogName","Message","TimeCreated","Count"'

function Write-Log([string]$msg) {
    Write-Host $msg
    Add-Content -Path $logFile -Value $msg -Encoding UTF8
}

# Collect events of one computer and write its CSV; $server = $null means local. Returns $true on success
function Export-ServerEvents([string]$server, [string]$name) {
    $remote = @{}
    if ($server) { $remote.ComputerName = $server }

    # Administrative logs with records (logs we cannot read are skipped; an unreachable computer is a failure)
    $listErr = $null
    try {
        $logs = @(Get-WinEvent @remote -ListLog * -ErrorAction SilentlyContinue -ErrorVariable listErr |
            Where-Object { $_.LogType -eq 'Administrative' -and $_.IsEnabled -and $_.RecordCount -gt 0 -and $_.LastWriteTime -gt $start })
    }
    catch {
        $listErr = @($_)
        $logs = @()
    }
    if ($logs.Count -eq 0 -and $listErr) {
        $reason = $listErr[0].Exception.Message
        if (-not $reason) { $reason = 'cannot connect (check host name, network/firewall, RPC and account permissions)' }
        Write-Log "  [ERROR] $name cannot list event logs: $reason"
        return $false
    }

    $events = foreach ($log in $logs) {
        try {
            Get-WinEvent @remote -FilterHashtable @{ LogName = $log.LogName; Level = 1, 2, 3; StartTime = $start } -ErrorAction Stop
        }
        catch {
            # "No matching events" is not an error (check ErrorId; the message text is localized)
            if ($_.FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*') {
                Write-Log "  [WARN] $name cannot read log [$($log.LogName)]: $($_.Exception.Message)"
            }
        }
    }

    # One row per provider + ID: TimeCreated = latest occurrence, Count = occurrences in the period
    $rows = @($events | Group-Object ProviderName, Id | ForEach-Object {
        $e = $_.Group | Sort-Object TimeCreated -Descending | Select-Object -First 1
        [pscustomobject]@{
            LevelDisplayName = $levelNames[[int]$e.Level]
            Id               = $e.Id
            LogName          = $e.LogName
            Message          = $e.Message
            TimeCreated      = $e.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            Count            = $_.Count
            Level            = [int]$e.Level
        }
    } | Sort-Object Level, Id | Select-Object LevelDisplayName, Id, LogName, Message, TimeCreated, Count)

    $csv = Join-Path $OutDir "${date}_$name.csv"
    if ($rows.Count -gt 0) {
        $rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    } else {
        # No errors/warnings: write a header-only CSV so the check sheet gets O (no CSV = collection failed)
        Set-Content -Path $csv -Value $csvHeader -Encoding UTF8
    }
    Write-Log "  $name : $($rows.Count) error/warning event types -> $csv"
    return $true
}

Write-Log ""
Write-Log "===== EventLogCheck_Local start $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') (last $Days days, computer: $env:COMPUTERNAME) ====="
if (Export-ServerEvents $null $env:COMPUTERNAME) {
    Write-Log "===== EventLogCheck_Local done ====="
    exit 0
}
Write-Log "===== EventLogCheck_Local FAILED ====="
exit 1
