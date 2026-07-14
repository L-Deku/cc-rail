@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0InstallPlugins.ps1" -SoftwareDir "%~dp0"
set "exit_code=%errorlevel%"
echo.
if not "%exit_code%"=="0" echo Installation failed with exit code %exit_code%.
pause
exit /b %exit_code%
