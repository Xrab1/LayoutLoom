@echo off
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
set "layoutloomExit=%errorlevel%"
if not "%layoutloomExit%"=="0" pause
exit /b %layoutloomExit%
