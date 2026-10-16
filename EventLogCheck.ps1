# 撈各伺服器近 7 天的事件檢視器紀錄，每台輸出一個 CSV：D:\SystemCheck\EventLog\日期_伺服器名.csv
$weekBefore = (Get-Date).AddDays(-7)
$Date = Get-Date -Format "yyyyMMdd"
$servers = Get-Content -Path D:\SystemCheck\serverlist.txt
foreach ($server in $servers) {
Get-WinEvent -ComputerName $server -ListLog * -EA silentlycontinue | Where-Object {$_.LogType -eq 'Administrative' -AND $_.isenabled -AND $_.lastwritetime -gt $weekBefore } |
Foreach-Object {
    Get-WinEvent -ComputerName $server -FilterHashTable @{LogName=$_.LogName; Level=1,2,3,4; StartTime=$WeekBefore} -EA SilentlyContinue|Sort-Object ProviderName, ID -unique
}|Sort-Object Level, Id|Select-Object -Property LevelDisplayName, Id, LogName, Message|Export-Csv -Path D:\SystemCheck\EventLog\$Date"_"$server.csv -NoTypeInformation -encoding utf8
}
