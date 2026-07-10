@echo off
title Wealth Pipeline App
cd /d "%~dp0"

echo Checking TradingView CDP connection...
curl -s http://localhost:9222/json/version >nul 2>&1
if %errorlevel% neq 0 (
    echo TradingView not detected on port 9222 - launching it...
    call scripts\launch_tv_debug.bat
)

start "" http://localhost:4590/
node scripts\pipeline_app_server.js
