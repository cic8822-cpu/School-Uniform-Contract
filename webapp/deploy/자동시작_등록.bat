@echo off
:: Uniform Purchase Guide - register auto-start on Windows login.
:: Delegates all logic to register-startup.ps1 (Korean text lives only in
:: the .ps1, which reads its own UTF-8 BOM correctly; cmd.exe does not).
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0register-startup.ps1"
pause
endlocal
