# 撈「本機」近 N 天的事件檢視器錯誤/警告，輸出一個 CSV：
#   <OutDir>\日期_電腦名稱.csv
# 用法（在每台伺服器上執行）：powershell -ExecutionPolicy Bypass -File EventLogCheck_Local.ps1 [-OutDir ...] [-LogDir ...] [-Days 7]
# - OutDir 可設成共用資料夾（例: \\fileserver\SystemCheck\EventLog），各台的 CSV 集中在一起給 fill_check 讀
# - 只抓 重大/錯誤/警告（Level 1~3），不抓資訊；沒有任何錯誤/警告時產生只有標題的 CSV
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

    # 有寫入紀錄的系統管理類記錄檔（個別記錄檔沒權限讀取時略過；整台連不上時視為失敗）
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
            # 沒有符合的事件不算錯誤（用 ErrorId 判斷，中文系統的訊息文字是中文）
            if ($_.FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*') {
                Write-Log "  [WARN] $name cannot read log [$($log.LogName)]: $($_.Exception.Message)"
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
