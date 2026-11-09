@echo off
cd /d "%~dp0"
if not exist .tmp mkdir .tmp
set "TEMP=%CD%\.tmp"
set "TMP=%CD%\.tmp"
py -3 -m venv .venv
if errorlevel 1 goto failed
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo Setup complete. Double-click Run JoeOS Write.cmd.
pause
exit /b
:failed
echo Setup failed. Install Python 3.12 or newer and check your internet connection.
pause
exit /b 1
