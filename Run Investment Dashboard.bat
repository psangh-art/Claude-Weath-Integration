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

REM Opens in a dedicated Chrome profile (user decision 2026-07-20) so a stale
REM dashboard tab from a previous launch can actually be CLOSED over CDP —
REM a browser refuses window.close() on a tab the user opened, so the web-only
REM guard can never do better than cover it. See scripts\dashboard_open.js.
start "Open Dashboard" /min node scripts\dashboard_open.js http://localhost:4600

REM Don't start a second dashboard server on 4600 — the running one already
REM serves the tab we just opened. (dashboard_server.js also refuses the port
REM gracefully, but this keeps the window from sitting on a dead process.)
echo Checking Investment Dashboard (port 4600)...
curl -s http://localhost:4600 >nul 2>&1
if %errorlevel% neq 0 (
    node scripts\dashboard_server.js
    pause
) else (
    echo Already running - reusing it. Closing this window.
)
