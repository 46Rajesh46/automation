@echo off
REM Uses %~dp0 = the folder this .bat file lives in — works on any drive automatically.
cd /d "%~dp0"
call venv\Scripts\activate.bat

echo [SCHEDULE] Setting yesterday dates in control.xlsx...
python set_yesterday_dates.py
if errorlevel 1 (
    echo [ERROR] set_yesterday_dates.py failed. Aborting.
    exit /b 1
)

echo [SCHEDULE] Starting core automation...
python core.py
