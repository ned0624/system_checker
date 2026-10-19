@echo off
REM Build exe files (requires: pip install pyinstaller pandas openpyxl)
REM Output goes to the dist\ folder
cd /d "%~dp0"

REM search_ignore: uses openpyxl only; exclude large unused packages for a small, fast exe
pyinstaller --onefile --exclude-module pandas --exclude-module numpy --exclude-module PIL --exclude-module lxml search_ignore.py

REM fill_check: needs pandas; --onedir avoids unpacking on every start
REM Output is the dist\fill_check\ folder - copy the whole folder and run fill_check.exe inside it
REM Put config.json in dist\fill_check\ next to fill_check.exe
pyinstaller --onedir --noconfirm --exclude-module PIL --exclude-module lxml fill_check.py

pause
