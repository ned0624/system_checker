@echo off
REM 週檢一鍵執行：1. 撈事件檢視器 CSV  2. 填檢核表、貼追蹤表
cd /d D:\SystemCheck

echo [1/2] 撈取事件檢視器紀錄...
powershell -ExecutionPolicy Bypass -File D:\SystemCheck\EventLogCheck.ps1

echo [2/2] 填寫檢核表與追蹤表...
REM 包成 exe 後改成: fill_check.exe
python D:\SystemCheck\fill_check.py

echo.
echo 完成，結果請看 D:\SystemCheck\logs
pause
