@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
 echo Jalankan install.ps1 terlebih dahulu.
 pause
 exit /b 1
)
".venv\Scripts\python.exe" main.py --watch
pause
