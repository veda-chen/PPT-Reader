@echo off
chcp 65001 >nul
title Smart PPT Reader - Launcher
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python not found. Install Python 3.10+ and check "Add Python to PATH".
  echo         https://www.python.org/downloads/
  pause
  exit /b 1
)

if not exist "backend\.deps_ok" (
  echo First run: installing dependencies, please wait ...
  python -m pip install -r "backend\requirements.txt"
  if errorlevel 1 (
    echo [ERROR] Failed to install dependencies. Check your network and retry.
    pause
    exit /b 1
  )
  echo ok> "backend\.deps_ok"
)

if not exist ".env" (
  if exist ".env.example" (
    copy /y .env.example .env >nul
    echo A new .env was created. Fill in your API Key in the Notepad window, save, then run this again.
    notepad ".env"
    exit /b 0
  )
)

echo Starting server ... browser will open at http://127.0.0.1:8800
echo (Server runs in a new window. Close that window to stop.)
cd /d "%~dp0backend"
start "Smart PPT Reader - server (close to stop)" python main.py
timeout /t 5 >nul
start "" http://127.0.0.1:8800
exit /b 0
