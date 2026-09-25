@echo off
:: Uniform Purchase Guide - remove auto-start shortcut.
:: Delegates all logic to unregister-startup.ps1.
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0unregister-startup.ps1"
pause
endlocal
