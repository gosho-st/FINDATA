@echo off
REM FINDATA — One-Click Installer for Windows

setlocal enabledelayedexpansion
cd /d "%~dp0"

cls
echo.
echo ╔════════════════════════════════════════════════════════════╗
echo ║    FINDATA — Financial Fundamental Data                    ║
echo ║    50,000+ tickers across 50+ global exchanges             ║
echo ╚════════════════════════════════════════════════════════════╝
echo.

REM Check Python
echo [*] Checking for Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Python not found! Please install from https://www.python.org/downloads/
    echo     Make sure to check "Add Python to PATH" during installation!
    pause
    start https://www.python.org/downloads/
    exit /b 1
)
echo [OK] Python found

REM Create venv
if not exist ".venv" (
    echo [*] Creating virtual environment...
    python -m venv .venv
    echo [OK] Virtual environment created
)

call .venv\Scripts\activate.bat

REM Install dependencies
echo [*] Checking dependencies...
python -c "import flask, pandas, selenium, xlsxwriter" >nul 2>&1
if %errorlevel% neq 0 (
    echo [*] Installing dependencies...
    pip install --upgrade pip --quiet
    pip install -r requirements.txt
    echo [OK] Dependencies installed
) else (
    echo [OK] All dependencies ready
)

REM Create desktop shortcut
set "DESKTOP=%USERPROFILE%\Desktop"
set "SHORTCUT=%DESKTOP%\FINDATA.lnk"
set "VBS=%TEMP%\shortcut.vbs"

if not exist "%SHORTCUT%" (
    echo [*] Creating desktop shortcut...
    echo Set ws = WScript.CreateObject("WScript.Shell") > "%VBS%"
    echo Set sc = ws.CreateShortcut("%SHORTCUT%") >> "%VBS%"
    echo sc.TargetPath = "%~dp0RUN_FINDATA.bat" >> "%VBS%"
    echo sc.WorkingDirectory = "%~dp0" >> "%VBS%"
    echo sc.Description = "FINDATA - Financial Data" >> "%VBS%"
    echo sc.Save >> "%VBS%"
    cscript //nologo "%VBS%"
    del "%VBS%"
    echo [OK] Desktop shortcut created
)

echo.
echo ════════════════════════════════════════════════════════════
echo    STARTING FINDATA...
echo    URL: http://127.0.0.1:5050
echo    Stop: Close this window or Ctrl+C
echo ════════════════════════════════════════════════════════════
echo.

python web_gui.py
pause
