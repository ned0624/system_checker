# Collect event log errors/warnings of the last N days on THIS computer, one CSV:
#   <OutDir>\yyyyMMdd_<computer name>.csv
# Usage (run on each server): powershell -ExecutionPolicy Bypass -File EventLogCheck_Local.ps1 [-OutDir ...] [-LogDir ...] [-Days 7] [-ExcludeFile ...]
# - OutDir can be a share (e.g. \\fileserver\SystemCheck\EventLog) to collect all servers' CSVs
# - Only Critical/Error/Warning (Level 1-3), no Information; header-only CSV when there are none
# - Event IDs listed in the exclude file are dropped
# - Exit code 1 when the CSV could not be written or the output/log folder is unusable
# - Apart from the SETTINGS block, keep this file pure ASCII
param(
    [string]$OutDir,
    [string]$LogDir,
    [int]$Days,
    [string]$ExcludeFile
)

# ===== SETTINGS =====================================================================
# Edit the values between the @' and '@ lines. Chinese folder names and network share
# paths (\\server\share\...) are OK: the script re-reads this block from the file itself and
# detects the file encoding (UTF-8 with/without BOM, or ANSI/Big5), so how Notepad saves
# the file does not matter. Command-line parameters (-OutDir etc.) override these values.
$Settings = @'
OutDir      = D:\SystemCheck\EventLog
LogDir      = D:\SystemCheck\logs
ExcludeFile = D:\SystemCheck\excluded_ids.txt
Days        = 7
'@
# ===== END OF SETTINGS ===============================================================

# Read the SETTINGS block from this script file with encoding detection.
# (PowerShell 5.1 itself reads a BOM-less file as ANSI, which garbles UTF-8 Chinese text.)
function Read-ScriptSettings {
    $bytes = [IO.File]::ReadAllBytes($PSCommandPath)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $text = [Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $text = [Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
    } else {
        try   { $text = (New-Object Text.UTF8Encoding($false, $true)).GetString($bytes) }   # strict UTF-8
        catch { $text = [Text.Encoding]::Default.GetString($bytes) }                         # ANSI (Big5 on zh-TW)
    }
    $result = @{}
    $m = [regex]::Match($text, '(?ms)^\$Settings\s*=\s*@''\r?\n(.*?)\r?\n''@')
    if (-not $m.Success) { return $result }
    foreach ($line in ($m.Groups[1].Value -split '\r?\n')) {
        $line = $line.Trim()
        if (-not $line -or $line.StartsWith('#')) { continue }
        $i = $line.IndexOf('=')
        if ($i -gt 0) { $result[$line.Substring(0, $i).Trim()] = $line.Substring($i + 1).Trim().Trim('"').Trim("'") }
    }
    return $result
}

# Parameters given on the command line win; otherwise use the SETTINGS block
$cfg = Read-ScriptSettings
foreach ($key in @('OutDir', 'LogDir', 'ExcludeFile', 'Days')) {
    if (-not $PSBoundParameters.ContainsKey($key)) {
        if (-not $cfg.ContainsKey($key) -or -not $cfg[$key]) {
            Write-Host "[ERROR] Setting '$key' is missing in the SETTINGS block of $PSCommandPath"
            exit 1
        }
        Set-Variable -Name $key -Value $cfg[$key]
    }
}
$Days = [int]$Days

$ErrorActionPreference = 'Continue'
$start   = (Get-Date).AddDays(-$Days)
$date    = Get-Date -Format 'yyyyMMdd'
$utf8Bom = New-Object Text.UTF8Encoding($true)

# Output and log folders must be usable, otherwise stop with exit code 1
foreach ($dir in @($OutDir, $LogDir)) {
    try {
        if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
            New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
        }
    }
    catch {
        Write-Host "[ERROR] Cannot create or access folder: $dir"
        Write-Host "        $($_.Exception.Message)"
        exit 1
    }
}
$logFile = Join-Path $LogDir "$date.log"
$logBroken = $false

# Level name from the level number, not LevelDisplayName (empty for some providers).
# Chinese names are built from char codes so this file stays pure ASCII:
#   1 = Critical (0x91CD 0x5927), 2 = Error (0x932F 0x8AA4), 3 = Warning (0x8B66 0x544A)
$levelNames = @{
    1 = [string][char]0x91CD + [char]0x5927
    2 = [string][char]0x932F + [char]0x8AA4
    3 = [string][char]0x8B66 + [char]0x544A
}
$csvHeader = '"LevelDisplayName","Id","LogName","Message","TimeCreated","Count"'

# Write to screen and log file (UTF-8). .NET file APIs are used so any valid path works.
function Write-Log([string]$msg) {
    Write-Host $msg
    try {
        [IO.File]::AppendAllText($logFile, $msg + "`r`n", $utf8Bom)
    }
    catch {
        if (-not $script:logBroken) {
            Write-Host "[WARN] Cannot write log file $logFile : $($_.Exception.Message)"
            $script:logBroken = $true
        }
    }
}

# Event IDs to drop before writing the CSV (one per line or comma separated, # = comment).
# Missing file = nothing excluded.
function Read-ExcludeIds {
    $ids = @{}
    if ($ExcludeFile -and (Test-Path -LiteralPath $ExcludeFile -PathType Leaf)) {
        Get-Content -LiteralPath $ExcludeFile | ForEach-Object { ($_ -split '#')[0] -split '[,\s]+' } |
            ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ } |
            ForEach-Object { $ids[$_] = $true }
        Write-Log "Excluded event IDs: $($ids.Count) (from $ExcludeFile)"
    } else {
        Write-Log "No exclude file ($ExcludeFile), no event IDs excluded"
    }
    return $ids
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

    # Drop excluded event IDs
    $all = @($events)
    $events = @($all | Where-Object { -not $excludeIds.ContainsKey([string]$_.Id) })
    $excludedCount = $all.Count - $events.Count

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

    # Both branches write through the PowerShell file system provider (works on DFS/UNC shares
    # where a direct .NET call failed). -Encoding UTF8 = UTF-8 with BOM in Windows PowerShell 5.1.
    $csv = Join-Path $OutDir "${date}_$name.csv"
    try {
        if ($rows.Count -gt 0) {
            $rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        } else {
            # No errors/warnings: header-only CSV so the check sheet gets O (no CSV = collection failed)
            $csvHeader | Out-File -LiteralPath $csv -Encoding UTF8 -ErrorAction Stop
        }
    }
    catch {
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        Write-Log "  [ERROR] $name cannot write CSV $csv"
        Write-Log "          $($ex.GetType().Name): $($ex.Message)"
        Write-Log "          folder exists: $(Test-Path -LiteralPath $OutDir -PathType Container), file exists: $(Test-Path -LiteralPath $csv), user: $env:USERDOMAIN\$env:USERNAME"
        return $false
    }
    Write-Log "  $name : $($rows.Count) error/warning event types -> $csv (excluded $excludedCount events)"
    return $true
}

Write-Log ""
Write-Log "===== EventLogCheck_Local start $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') (last $Days days, computer: $env:COMPUTERNAME) ====="
Write-Log "CSV folder : $OutDir"
$excludeIds = Read-ExcludeIds
if (Export-ServerEvents $null $env:COMPUTERNAME) {
    Write-Log "===== EventLogCheck_Local done ====="
    exit 0
}
Write-Log "===== EventLogCheck_Local FAILED ====="
exit 1
