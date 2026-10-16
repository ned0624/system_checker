@echo off
REM 週檢一鍵執行：1. 撈事件檢視器 CSV  2. 填檢核表、貼追蹤表
REM 可直接雙擊執行，也可排進工作排程器
set BASE=D:\SystemCheck

for /f %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set TODAY=%%d

echo [1/2] 撈取事件檢視器紀錄...
powershell -NoProfile -ExecutionPolicy Bypass -File "%BASE%\EventLogCheck.ps1"

REM 沒撈到今天的 CSV 就停止，避免 fill_check 拿上週的舊 CSV 重跑、蓋掉上週的追蹤表
if not exist "%BASE%\EventLog\%TODAY%_*.csv" (
    echo [錯誤] 沒有產生今天 %TODAY% 的 CSV，停止執行，請確認伺服器連線與 serverlist.txt
    goto END
)

echo [2/2] 填寫檢核表與追蹤表...
REM 有 fill_check.exe 就用 exe，沒有就用 Python 執行 fill_check.py
if exist "%BASE%\fill_check.exe" (
    "%BASE%\fill_check.exe"
) else (
    where python >nul 2>nul
    if errorlevel 1 (
        echo [錯誤] 找不到 %BASE%\fill_check.exe，這台電腦也沒有安裝 Python
        goto END
    )
    python "%BASE%\fill_check.py"
)
if errorlevel 1 (
    echo [錯誤] 填表沒有完成，請看 %BASE%\logs 的 log
) else (
    echo 完成，結果請看 %BASE%\logs
)

:END
echo.
pause
