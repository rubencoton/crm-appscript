@echo off
set ROOT=C:\Users\elrub\Desktop\CARPETA CODEX
powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT%\06_AUTOMATIZACION\sync-all.ps1" -Mode backup -WorkspaceRoot "%ROOT%"
pause
