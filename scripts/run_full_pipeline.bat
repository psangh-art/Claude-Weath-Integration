@echo off
title TradingView Full Pipeline
cd /d "%~dp0\.."

echo Checking TradingView CDP connection...
curl -s http://localhost:9222/json/version >nul 2>&1
if %errorlevel% neq 0 (
    echo TradingView not detected on port 9222 - launching it...
    call scripts\launch_tv_debug.bat
)

REM Runs the full periodic pipeline: chart capture -> tradingview_layouts.xlsx
REM -> OCR channel detection -> Stocks_Buy_Strategy.xlsx update -> feedback log.
REM Degrades gracefully to chart-export-only if Tesseract OCR or the master
REM sheet aren't present yet — see run_full_pipeline.js.
node scripts\run_full_pipeline.js

echo.
pause
