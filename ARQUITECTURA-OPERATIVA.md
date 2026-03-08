# ARQUITECTURA-OPERATIVA

## Ruta canonica
- `C:\Users\elrub\Desktop\CARPETA CODEX\01_PROYECTOS\festivales-github`

## Flujo diario
1. Abrir workspace canonico: `C:\Users\elrub\Desktop\CARPETA CODEX`
2. Ejecutar `sync_in` (manual o por tarea)
3. Trabajar en codigo
4. Ejecutar `sync_out`
5. Verificar `git status -sb` limpio

## Capas
- `sync_in.ps1`: entrada manual estricta (git pull --rebase + clasp pull en 2 proyectos).
- `sync_out.ps1`: salida manual (clasp push en 2 proyectos + git add/commit/push).
- `sync_full.ps1`: ciclo completo manual.
- `sync_in_auto.ps1`: entrada automatica segura (cada 5 min), con lock y logs.

## Automatizacion
- Tarea Windows: `CodexSyncInFestivales_Every5Min` -> `sync_in_auto.ps1`.
- Fallback de inicio de sesion: Startup `CodexSyncInOnLogon.bat` -> `sync_in_auto.ps1`.

## Gobernanza de conflictos
- Nunca forzar `push` ni reescrituras destructivas.
- Si `git pull --rebase` da conflicto: resolver manualmente o `git rebase --abort`.

## Sincronizacion entre dispositivos
- Codigo: GitHub + scripts `sync_*`.
- Hilos Codex: misma cuenta + mismo workspace + cloud setup completado.
- Diagnostico/alineacion en `codex-tools`.
