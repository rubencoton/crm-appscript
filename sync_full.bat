@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0sync_full.ps1" %*
if errorlevel 1 (
  echo ERROR en sync_full
)
pause
