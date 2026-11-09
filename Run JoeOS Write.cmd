@echo off
cd /d "%~dp0"
if exist "dist-keyboard\JoeOS Write\JoeOS Write.exe" (
  start "" "dist-keyboard\JoeOS Write\JoeOS Write.exe" %*
  exit /b
)
if exist ".venv\Scripts\pythonw.exe" (
  start "" ".venv\Scripts\pythonw.exe" joeos_app.py %*
  exit /b
)
echo Run Setup.cmd first, then try again.
pause
