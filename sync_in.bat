@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0sync_in.ps1" %*
if errorlevel 1 (
  echo ERROR en sync_in
)
pause
