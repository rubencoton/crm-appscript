# README-SYNC

## Objetivo
Mantener este repo sincronizado entre PC y portatil con GitHub + Google Apps Script.

## Ruta canonica
- `C:\Users\elrub\Desktop\CARPETA CODEX\01_PROYECTOS\festivales-github`

## Scripts principales
- Entrada: `sync_in.ps1` / `sync_in.bat`
- Salida: `sync_out.ps1` / `sync_out.bat`
- Flujo completo: `sync_full.ps1` / `sync_full.bat`

## Uso recomendado diario
1. `sync_in.bat`
2. Editar y probar
3. `sync_out.bat`

## Modo terminal
- `powershell -NoProfile -ExecutionPolicy Bypass -File .\sync_in.ps1`
- `powershell -NoProfile -ExecutionPolicy Bypass -File .\sync_out.ps1`
- `powershell -NoProfile -ExecutionPolicy Bypass -File .\sync_full.ps1`

## Modo prueba (sin tocar remoto)
- `powershell -NoProfile -ExecutionPolicy Bypass -File .\sync_in.ps1 -DryRun -AllowDirty`
- `powershell -NoProfile -ExecutionPolicy Bypass -File .\sync_out.ps1 -DryRun`
- `powershell -NoProfile -ExecutionPolicy Bypass -File .\sync_full.ps1 -DryRun -AllowDirty`

## Politica de seguridad
- Sin comandos destructivos.
- `sync_in` usa `git pull --ff-only` (sin rebase forzado).
- `sync_in` bloquea si hay cambios locales, salvo `-AllowDirty`.
- `sync_out` bloquea volumen alto de no-trackeados salvo `-AllowBulkUntracked`.
- Si hay conflicto Git, se para el flujo para resolver manualmente.

## Automatizacion
- Tarea programada 5 min: `CodexSyncIn_5min`
- Trigger ONLOGON: si no hay permisos, fallback en Startup.
- Fallback actual: `C:\Users\elrub\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\CodexSyncIn_OnLogon.cmd`

## Verificaciones rapidas
- Estado Git: `git status -sb`
- Estado clasp: `npx clasp status`
- Estado hilos Codex: `powershell -NoProfile -ExecutionPolicy Bypass -File .\codex-tools\diagnostico_hilos_codex.ps1`
