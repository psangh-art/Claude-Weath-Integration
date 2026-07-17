@echo off
cd /d "%~dp0"
start "" http://localhost:4600
node scripts\dashboard_server.js
pause
