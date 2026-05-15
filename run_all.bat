@echo off
cd /d "%~dp0"
call venv\Scripts\activate.bat

echo [SCHEDULE] Setting yesterday dates in control.xlsx...
python set_yesterday_dates.py
if errorlevel 1 (
    echo [ERROR] set_yesterday_dates.py failed. Aborting.
    pause
    exit /b 1
)

echo [SCHEDULE] Starting core automation...
python core.py

echo.
echo ==========================================
echo   run_all.bat finished. Check above for errors.
echo ==========================================
pause
