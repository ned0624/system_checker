# 撈 serverlist.txt 中每台伺服器近 N 天的事件檢視器錯誤/警告，每台輸出一個 CSV：
#   <OutDir>\日期_伺服器名.csv
# 用法：powershell -ExecutionPolicy Bypass -File EventLogCheck.ps1 [-ServerList ...] [-OutDir ...] [-LogDir ...] [-Days 7]
# - 只抓 重大/錯誤/警告（Level 1~3），不抓資訊
# - 連不上的伺服器不產生 CSV，並寫入 log；沒有任何錯誤/警告的伺服器會產生只有標題的 CSV
# - 執行帳號需能遠端讀取各伺服器的事件記錄
param(
    [string]$ServerList = 'D:\SystemCheck\serverlist.txt',
    [string]$OutDir     = 'D:\SystemCheck\EventLog',
    [string]$LogDir     = 'D:\SystemCheck\logs',
    [int]$Days          = 7
)

$ErrorActionPreference = 'Continue'
$start = (Get-Date).AddDays(-$Days)
$date  = Get-Date -Format 'yyyyMMdd'
New-Item -ItemType Directory -Force -Path $OutDir, $LogDir | Out-Null
$logFile = Join-Path $LogDir "$date.log"

# 等級用數字對應，不靠 LevelDisplayName（部分事件來源沒有顯示名稱會是空白）
$levelNames = @{ 1 = '重大'; 2 = '錯誤'; 3 = '警告' }
$csvHeader  = '"LevelDisplayName","Id","LogName","Message","TimeCreated","Count"'

function Write-Log([string]$msg) {
    Write-Host $msg
    Add-Content -Path $logFile -Value $msg -Encoding UTF8
}

# 撈一台的事件並輸出 CSV；$server 為 $null 時撈本機。成功回傳 $true
function Export-ServerEvents([string]$server, [string]$name) {
    $remote = @{}
    if ($server) { $remote.ComputerName = $server }

    # 有寫入紀錄的系統管理類記錄檔
    $listErr = $null
    $logs = @(Get-WinEvent @remote -ListLog * -ErrorAction SilentlyContinue -ErrorVariable listErr |
        Where-Object { $_.LogType -eq 'Administrative' -and $_.IsEnabled -and $_.RecordCount -gt 0 -and $_.LastWriteTime -gt $start })
    if ($logs.Count -eq 0 -and $listErr) {
        Write-Log "  [錯誤] $name 無法讀取記錄檔清單（連線或權限問題）: $($listErr[0].Exception.Message)"
        return $false
    }

    $events = foreach ($log in $logs) {
        try {
            Get-WinEvent @remote -FilterHashtable @{ LogName = $log.LogName; Level = 1, 2, 3; StartTime = $start } -ErrorAction Stop
        }
        catch {
            # 沒有符合的事件不算錯誤（用 ErrorId 判斷，中文系統的訊息文字是中文）
            if ($_.FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*') {
                Write-Log "  [警告] $name 記錄檔 [$($log.LogName)] 讀取失敗: $($_.Exception.Message)"
            }
        }
    }

    # 同一來源、同一 ID 只留一筆：時間取最近一次，Count 為期間內發生次數
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
        # 沒有錯誤/警告也輸出只有標題的 CSV，檢核表才會填 O（沒有 CSV 代表撈取失敗）
        Set-Content -Path $csv -Value $csvHeader -Encoding UTF8
    }
    Write-Log "  $name：錯誤/警告 $($rows.Count) 種 → $csv"
    return $true
}

Write-Log ""
Write-Log "===== EventLogCheck 開始 $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')（近 $Days 天，伺服器清單: $ServerList）====="
if (-not (Test-Path $ServerList)) {
    Write-Log "[錯誤] 找不到伺服器清單: $ServerList"
    exit 1
}
# 一行一台，空行與 # 開頭的行略過
$servers = @(Get-Content $ServerList | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') })

$failed = @()
foreach ($s in $servers) {
    Write-Log $s
    if (-not (Export-ServerEvents $s $s)) { $failed += $s }
}
Write-Log "===== EventLogCheck 完成：伺服器 $($servers.Count) 台，失敗 $($failed.Count) 台 ====="
foreach ($s in $failed) { Write-Log "  撈取失敗: $s" }
exit 0
