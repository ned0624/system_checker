# Collect event log errors/warnings of a period (default: last 7 days) on THIS computer, one CSV:
#   <OutDir>\yyyyMMdd_<computer name>.csv
# Usage (run on each server): powershell -ExecutionPolicy Bypass -File EventLogCheck_Local.ps1 [-OutDir ...] [-LogDir ...] [-StartTime 2026-10-15] [-EndTime 2026-10-22] [-Days 7] [-ExcludeFile ...]
# - OutDir can be a share (e.g. \\fileserver\SystemCheck\EventLog) to collect all servers' CSVs
# - Only Critical/Error/Warning (Level 1-3), no Information; no CSV when none are left after exclusion
# - Event IDs listed in the exclude file are dropped
# - Exit code: 0 = CSV written, 1 = error (CSV/folder not writable), 2 = nothing to report
# - Apart from the SETTINGS block, keep this file pure ASCII
param(
    [string]$OutDir,
    [string]$LogDir,
    [string]$StartTime,
    [string]$EndTime,
    [string]$Days,
    [string]$ExcludeFile
)

# ===== SETTINGS =====================================================================
# Edit the values between the @' and '@ lines. Chinese folder names and network share
# paths (\\server\share\...) are OK: the script re-reads this block from the file itself and
# detects the file encoding (UTF-8 with/without BOM, or ANSI/Big5), so how Notepad saves
# the file does not matter. Command-line parameters (-OutDir etc.) override these values.
# Period: StartTime / EndTime as yyyy-MM-dd or yyyy-MM-dd HH:mm (date only = 00:00). Empty EndTime = now,
# empty StartTime = EndTime minus Days, empty Days = 7.
# Paths may contain {week} (W2643), {yy} (26), {ww} (43) - replaced with the ISO week of the period end,
# e.g. OutDir = Z:\...\System Check\{week}\Reference\BOS
$Settings = @'
OutDir      = D:\SystemCheck\EventLog
LogDir      = D:\SystemCheck\logs
ExcludeFile = D:\SystemCheck\excluded_ids.txt
StartTime   =
EndTime     =
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
foreach ($key in @('OutDir', 'LogDir', 'ExcludeFile')) {
    if (-not $PSBoundParameters.ContainsKey($key)) {
        if (-not $cfg.ContainsKey($key) -or -not $cfg[$key]) {
            Write-Host "[ERROR] Setting '$key' is missing in the SETTINGS block of $PSCommandPath"
            exit 1
        }
        Set-Variable -Name $key -Value $cfg[$key]
    }
}
# Optional: StartTime / EndTime / Days (empty = not set)
foreach ($key in @('StartTime', 'EndTime', 'Days')) {
    if (-not $PSBoundParameters.ContainsKey($key) -and $cfg.ContainsKey($key)) {
        Set-Variable -Name $key -Value $cfg[$key]
    }
}

# Parse a date/time setting; date only = 00:00 of that day. Fixed formats, independent of regional settings.
function ConvertTo-Time([string]$text, [string]$name) {
    $formats = [string[]]@('yyyy-MM-dd', 'yyyy-MM-dd HH:mm', 'yyyy-MM-dd HH:mm:ss', 'yyyy-M-d', 'yyyy-M-d H:mm',
                           'yyyy/MM/dd', 'yyyy/MM/dd HH:mm', 'yyyy/MM/dd HH:mm:ss', 'yyyy/M/d', 'yyyy/M/d H:mm')
    try {
        return [datetime]::ParseExact($text.Trim(), $formats, [Globalization.CultureInfo]::InvariantCulture,
                                      [Globalization.DateTimeStyles]::None)
    }
    catch {
        Write-Host "[ERROR] $name '$text' is not a valid date/time (use yyyy-MM-dd or yyyy-MM-dd HH:mm)"
        exit 1
    }
}

if ([string]::IsNullOrWhiteSpace($Days)) {
    $Days = 7
} elseif (-not [int]::TryParse($Days.Trim(), [ref]$null) -or [int]$Days -lt 0) {
    Write-Host "[ERROR] Days '$Days' must be a whole number of days"
    exit 1
}
$Days = [int]$Days

# ISO week label, e.g. 2026-10-22 -> W2643 (same rule as fill_check / Python isocalendar).
# .NET Framework has no ISOWeek and GetWeekOfYear differs around New Year, so compute it:
# the ISO week/year are those of the Thursday in the same Monday-Sunday week.
function Get-IsoWeek([datetime]$d) {
    $thu  = $d.Date.AddDays(3 - (([int]$d.DayOfWeek + 6) % 7))
    $week = [math]::Floor(($thu.DayOfYear - 1) / 7) + 1
    return 'W{0:D2}{1:D2}' -f ($thu.Year % 100), [int]$week
}

$ErrorActionPreference = 'Continue'
# Period: EndTime (default now) back to StartTime (default EndTime - Days, Days default 7)
$end = if ([string]::IsNullOrWhiteSpace($EndTime)) { Get-Date } else { ConvertTo-Time $EndTime 'EndTime' }
$start = if ([string]::IsNullOrWhiteSpace($StartTime)) { $end.AddDays(-$Days) } else { ConvertTo-Time $StartTime 'StartTime' }
if ($start -ge $end) {
    Write-Host "[ERROR] StartTime ($($start.ToString('yyyy-MM-dd HH:mm'))) must be earlier than EndTime ($($end.ToString('yyyy-MM-dd HH:mm')))"
    exit 1
}
# CSV file date and {week} follow the END of the period (= today when EndTime is not set)
$date    = $end.ToString('yyyyMMdd')
$week    = Get-IsoWeek $end
$period  = "$($start.ToString('yyyy-MM-dd HH:mm')) ~ $($end.ToString('yyyy-MM-dd HH:mm')) (week $week)"

# {week} / {yy} / {ww} in paths -> this week, e.g. ...\System Check\{week}\Reference\BOS -> ...\W2643\...
foreach ($key in @('OutDir', 'LogDir', 'ExcludeFile')) {
    $val = Get-Variable -Name $key -ValueOnly
    if ($val) {
        Set-Variable -Name $key -Value ($val.Replace('{week}', $week).Replace('{yy}', $week.Substring(1, 2)).Replace('{ww}', $week.Substring(3, 2)))
    }
}
$utf8Bom = New-Object Text.UTF8Encoding($true)

# Output and log folders must exist and be writable by this account, otherwise stop with exit code 1.
# (On a share with access-based enumeration a folder without permission looks like it does not exist.)
$account = "$env:USERDOMAIN\$env:USERNAME"
foreach ($dir in @($OutDir, $LogDir)) {
    try {
        if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
            New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
        }
        if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
            throw "folder is not visible to this account (no permission, or the path is wrong)"
        }
        $probe = Join-Path $dir ('.write_test_' + [guid]::NewGuid().ToString('N') + '.tmp')
        'test' | Out-File -LiteralPath $probe -Encoding ASCII -ErrorAction Stop
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
    }
    catch {
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        Write-Host "[ERROR] Cannot write to folder: $dir"
        Write-Host "        $($ex.Message)"
        Write-Host "        Account $account needs Modify permission on this folder (share and NTFS),"
        Write-Host "        or run this script as an account that has it."
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

# Collect events of one computer and write its CSV; $server = $null means local.
# Returns 'ok' (CSV written), 'empty' (no events left after exclusion, no CSV) or 'fail'.
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
        return 'fail'
    }

    $events = foreach ($log in $logs) {
        try {
            Get-WinEvent @remote -FilterHashtable @{ LogName = $log.LogName; Level = 1, 2, 3; StartTime = $start; EndTime = $end } -ErrorAction Stop
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

    # Nothing left after exclusion: no CSV, only a log line
    if ($rows.Count -eq 0) {
        Write-Log "  $name : no error/warning events (excluded $excludedCount events), no CSV written"
        return 'empty'
    }

    # -Encoding UTF8 = UTF-8 with BOM in Windows PowerShell 5.1
    $csv = Join-Path $OutDir "${date}_$name.csv"
    try {
        $rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        Write-Log "  [ERROR] $name cannot write CSV $csv"
        Write-Log "          $($ex.GetType().Name): $($ex.Message)"
        Write-Log "          folder visible: $(Test-Path -LiteralPath $OutDir -PathType Container), account: $account"
        return 'fail'
    }
    Write-Log "  $name : $($rows.Count) error/warning event types -> $csv (excluded $excludedCount events)"
    return 'ok'
}

Write-Log ""
Write-Log "===== EventLogCheck_Local start $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') (computer: $env:COMPUTERNAME) ====="
Write-Log "Period     : $period"
Write-Log "CSV folder : $OutDir"
$excludeIds = Read-ExcludeIds
switch (Export-ServerEvents $null $env:COMPUTERNAME) {
    'ok'    { Write-Log "===== EventLogCheck_Local done ====="; exit 0 }
    # Exit code 2 = nothing to fill in (run_weekly.bat skips fill_check)
    'empty' { Write-Log "===== EventLogCheck_Local done, no CSV written ====="; exit 2 }
    default { Write-Log "===== EventLogCheck_Local FAILED ====="; exit 1 }
}
