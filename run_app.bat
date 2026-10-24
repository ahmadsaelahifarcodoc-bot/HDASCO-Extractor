@echo off
setlocal
cd /d "%~dp0"
py -m pip install -r requirements.txt
if errorlevel 1 (
  echo Package installation failed. Check Python and internet access.
  pause
  exit /b 1
)
py main.py
pause
