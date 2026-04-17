@echo off
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File "%~dp0instalar_windows.ps1"
pause
