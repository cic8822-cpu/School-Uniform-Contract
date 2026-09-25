@echo off
:: Uniform Purchase Guide - double-click launcher.
:: Starts server.ps1 (hidden) and opens the default browser.
:: Details (Korean): see DEPLOY.md in this project.
chcp 65001 >nul
setlocal

set "SCRIPT_DIR=%~dp0"
set "PORT=8973"

start "" /min powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%SCRIPT_DIR%server.ps1" -Port %PORT%

timeout /t 2 /nobreak >nul

start "" "http://localhost:%PORT%/"

endlocal
