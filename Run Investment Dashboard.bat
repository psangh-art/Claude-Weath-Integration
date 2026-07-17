@echo off
title Investment Dashboard
cd /d "%~dp0"

REM The dashboard is the main screen now — this is the one shortcut the user
REM needs. It also brings up the Investment Production Centre (port 4590) in
REM the background so the dashboard's "Pipeline" nav link works immediately,
REM instead of requiring "Run Pipeline App.bat" to be started separately.
echo Checking Investment Production Centre (port 4590)...
curl -s http://localhost:4590 >nul 2>&1
if %errorlevel% neq 0 (
    echo Not running - starting it in the background...
    start "Investment Production Centre" /min node scripts\pipeline_app_server.js
) else (
    echo Already running.
)

start "" http://localhost:4600
node scripts\dashboard_server.js
pause
