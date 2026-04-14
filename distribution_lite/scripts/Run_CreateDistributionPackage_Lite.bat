@echo off
setlocal
set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%CreateDistributionPackage_Lite.ps1" %*
if errorlevel 1 exit /b %errorlevel%
pause
