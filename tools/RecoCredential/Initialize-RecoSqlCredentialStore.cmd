@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Initialize-RecoSqlCredentialStore.ps1"
set "exit_code=%errorlevel%"
echo.
if not "%exit_code%"=="0" echo SQL credential initialization failed with exit code %exit_code%.
pause
exit /b %exit_code%
