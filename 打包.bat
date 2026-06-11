@echo off
chcp 65001 >nul
title Build - Smart PPT Reader
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python not found. Please install Python 3.10+ first.
  echo         https://www.python.org/downloads/
  pause
  exit /b 1
)

echo [1/3] Installing runtime dependencies ...
python -m pip install -r "backend\requirements.txt"
if errorlevel 1 (
  echo [ERROR] Failed to install dependencies.
  pause
  exit /b 1
)

echo [2/3] Installing PyInstaller ...
python -m pip install pyinstaller
if errorlevel 1 (
  echo [ERROR] Failed to install PyInstaller.
  pause
  exit /b 1
)

echo [3/3] Building exe (first build is slow, please wait) ...
python -m PyInstaller build.spec --noconfirm --clean
if errorlevel 1 (
  echo [ERROR] Build failed. Please copy the error above and send it to me.
  pause
  exit /b 1
)

echo Copying config template and readme into dist ...
copy /y .env.example dist >nul
copy /y *.txt dist >nul

echo.
echo ====================================================
echo  DONE. The exe is in the "dist" folder.
echo  Send the whole "dist" folder to others.
echo ====================================================
pause
