@echo off
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\clear-office-cache.ps1" %*
echo.
pause
