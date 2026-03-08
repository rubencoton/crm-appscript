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

## Uso automatico
- Cada 5 minutos: tarea `CodexSyncInFestivales_Every5Min` ejecuta `sync_in_auto.ps1`.
- Al iniciar sesion: fallback en Startup (`CodexSyncInOnLogon.bat`).

## Salud del sistema
- Ejecutar `sync_healthcheck.ps1` para ver estado integral (Git/GAS/Codex/tareas).
- Reporte generado en `logs\healthcheck-latest.txt`.

## Politica de seguridad Git
- Sin comandos destructivos.
- Si hay conflicto/rebase bloqueado, no forzar.
- Resolver manualmente o usar `git rebase --abort`.

## Verificaciones rapidas
- Estado Git: `git status -sb`
- Estado clasp: `clasp status`
- Estado hilos Codex: `powershell -NoProfile -ExecutionPolicy Bypass -File .\codex-tools\diagnostico_hilos_codex.ps1`
