@echo off
rem Double-click me to validate RedisExcel.json against the shipped schema.
rem Keeps the window open so you can read the result.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\validate-config.ps1" %*
echo.
pause
