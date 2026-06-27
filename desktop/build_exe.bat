@echo off
REM ============================================================
REM   Meter CRM  -  one-click builder for Windows (.exe)
REM   Double-click this file. It builds dist\MeterCRM.exe
REM ============================================================
setlocal
cd /d "%~dp0"

echo.
echo  [1/3] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
  echo.
  echo  Python was not found. Please install Python 3 from:
  echo      https://www.python.org/downloads/windows/
  echo  During install, TICK "Add python.exe to PATH", then run this again.
  echo.
  pause
  exit /b 1
)

echo  [2/3] Installing build tools (PyInstaller + openpyxl)...
python -m pip install --upgrade pip >nul
python -m pip install pyinstaller openpyxl
if errorlevel 1 (
  echo  Could not install build tools. Check your internet connection.
  pause
  exit /b 1
)

echo  [3/3] Building MeterCRM.exe ...
python -m PyInstaller --noconfirm --onefile --windowed --name MeterCRM ^
  --icon app.ico --add-data "app.ico;." meter_crm.py
if errorlevel 1 (
  echo  Build failed. See messages above.
  pause
  exit /b 1
)

copy /Y meter_data.json dist\meter_data.json >nul

echo.
echo  ============================================================
echo   DONE!  Your app is here:   dist\MeterCRM.exe
echo   Keep MeterCRM.exe and meter_data.json together in one folder.
echo   (You can copy that folder anywhere - Desktop, USB, etc.)
echo  ============================================================
echo.
start "" "%cd%\dist"
pause
