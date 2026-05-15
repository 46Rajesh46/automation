@echo off
cd /d "%~dp0"
call venv\Scripts\activate.bat
python send_zonal_dashboard_email.py
