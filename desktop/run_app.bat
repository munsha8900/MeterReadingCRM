@echo off
REM Run Meter CRM directly with Python (no build needed).
setlocal
cd /d "%~dp0"
python --version >nul 2>&1 || (echo Install Python 3 first: https://www.python.org/downloads/windows/ & pause & exit /b 1)
python -m pip install openpyxl >nul 2>&1
python meter_crm.py
