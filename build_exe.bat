@echo off
setlocal
cd /d "%~dp0"
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
py -m PyInstaller --noconfirm --clean --onefile --windowed --name HDASCO_Extractor main.py
echo.
echo Finished. EXE should be in the dist folder.
pause
