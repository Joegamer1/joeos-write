@echo off
cd /d "%~dp0"
set "TEMP=%CD%\.tmp"
set "TMP=%CD%\.tmp"
set "PYINSTALLER_CONFIG_DIR=%CD%\.build-cache"
.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --distpath dist-keyboard "JoeOS Write.spec"
if errorlevel 1 exit /b 1
echo Portable app: dist-keyboard\JoeOS Write\JoeOS Write.exe
