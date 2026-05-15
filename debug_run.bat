@echo off
cd /d "%~dp0"
call venv\Scripts\activate.bat

echo ===== DEBUG RUN %date% %time% ===== > debug_log.txt

echo [1] Testing Python imports... >> debug_log.txt
python -c "import pandas; print('pandas OK')" >> debug_log.txt 2>&1
python -c "import selenium; print('selenium OK')" >> debug_log.txt 2>&1
python -c "import win32com; print('win32com OK')" >> debug_log.txt 2>&1
python -c "import openpyxl; print('openpyxl OK')" >> debug_log.txt 2>&1
python -c "import plotly; print('plotly OK')" >> debug_log.txt 2>&1
python -c "import pywinauto; print('pywinauto OK')" >> debug_log.txt 2>&1
python -c "import xlrd; print('xlrd OK')" >> debug_log.txt 2>&1

echo. >> debug_log.txt
echo [2] Running set_yesterday_dates.py... >> debug_log.txt
python set_yesterday_dates.py >> debug_log.txt 2>&1
echo Exit code: %errorlevel% >> debug_log.txt

echo. >> debug_log.txt
echo [3] Running core.py... >> debug_log.txt
python core.py >> debug_log.txt 2>&1
echo Exit code: %errorlevel% >> debug_log.txt

echo. >> debug_log.txt
echo ===== DONE ===== >> debug_log.txt

echo Finished. Open debug_log.txt in Notepad to see what happened.
notepad debug_log.txt
