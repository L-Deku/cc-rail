@echo off
setlocal
rem Re-grant only the fixed databases (RecoLearning / RecoData2024 / RecoData2020 / model) on both servers,
rem skip the per-project-database loop, then regenerate RecoPluginSql.json. Takes about one minute.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Invoke-RecoPluginLoginSetup.ps1" -SkipProjectDatabases
set "exit_code=%errorlevel%"
echo.
if not "%exit_code%"=="0" echo Setup failed with exit code %exit_code%. Please send the red text above to Claude.
pause
exit /b %exit_code%
