@echo off
REM FINDATA — Quick Launcher for Windows
cd /d "%~dp0"

REM Check if venv exists
if not exist ".venv" (
    echo First run - setting up environment...
    python -m venv .venv
    call .venv\Scripts\activate.bat
    pip install --upgrade pip --quiet
    pip install -r requirements.txt
    echo Setup complete!
) else (
    call .venv\Scripts\activate.bat
)

echo.
echo ════════════════════════════════════════════════════════════
echo   FINDATA is running!
echo   URL: http://127.0.0.1:5050
echo   To stop: Close this window
echo ════════════════════════════════════════════════════════════
echo.

start http://127.0.0.1:5050
python web_gui.py
pause
