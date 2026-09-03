@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Invoke-RecoPluginLoginSetup.ps1"
set "exit_code=%errorlevel%"
echo.
if not "%exit_code%"=="0" echo Setup failed with exit code %exit_code%. Please send the red text above to Claude.
pause
exit /b %exit_code%
