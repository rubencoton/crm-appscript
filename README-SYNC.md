# README-SYNC

## Flujo rapido recomendado
1. Traer cambios antes de editar:
- `sync_in.bat`

2. Subir cambios al terminar:
- `sync_out.bat`

3. Flujo completo:
- `sync_full.bat`

## Modo auditor (forzar chequeos)
- `npm run sync:audit`

Genera un reporte en:
- `C:\Users\elrub\Desktop\CARPETA CODEX\05_DOCUMENTACION\AUDITORIA_SYNC_*.md`

## Modo test sin tocar remoto (DryRun)
- `powershell -ExecutionPolicy Bypass -File .\sync_in.ps1 -DryRun`
- `powershell -ExecutionPolicy Bypass -File .\sync_out.ps1 -DryRun`
- `powershell -ExecutionPolicy Bypass -File .\sync_full.ps1 -DryRun`

## Terminal (sin .bat)
- Entrada: `powershell -ExecutionPolicy Bypass -File .\sync_in.ps1`
- Salida: `powershell -ExecutionPolicy Bypass -File .\sync_out.ps1`
- Completo: `powershell -ExecutionPolicy Bypass -File .\sync_full.ps1`

## Seguridad
- No se hace `--force` ni en git ni en clasp.
- Si hay conflicto git/rebase, se para el flujo.
- `sync_in` bloquea por defecto si hay cambios locales (usar `-AllowDirty` solo si sabes lo que haces).
- `sync_out` admite `-SkipGitPush` para validar sin publicar.

## Regla operativa
- Usa siempre el mismo workspace en ambos equipos:
  - `C:\Users\elrub\Desktop\CARPETA CODEX`
