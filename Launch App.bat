@echo off
title Select by G — Supplier Presentation Generator
color 0F

echo.
echo  ============================================================
echo    Select by G Group - Supplier Presentation Generator
echo  ============================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python is not installed or not in PATH.
    echo.
    echo  Please install Python from https://www.python.org/downloads/
    echo  Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)
echo  [OK] Python found.

cd /d "%~dp0"

echo  [..] Installing dependencies...
python -m pip install -q flask lxml Pillow 2>nul
echo  [OK] Dependencies ready.

for /f "tokens=5" %%a in ('netstat -aon 2^>nul ^| findstr ":5001 "') do (
    taskkill /PID %%a /F >nul 2>&1
)

echo  [..] Starting server...
start "" /MIN python app.py

echo  [..] Waiting for server (5 seconds)...
timeout /t 5 /nobreak >nul

echo  [OK] Opening browser at http://localhost:5001
start http://localhost:5001

echo.
echo  ============================================================
echo    App is running in the background.
echo    You can close this window - the server keeps running.
echo    To stop: close the minimized Python window in taskbar.
echo  ============================================================
echo.
pause
