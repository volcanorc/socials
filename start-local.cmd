@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -NoExit -ExecutionPolicy Bypass -File "%~dp0run-server.ps1"
endlocal
