# ARQUITECTURA-OPERATIVA

## Flujo diario
1. Abrir workspace canonico: `C:\Users\elrub\Desktop\CARPETA CODEX`
2. Ejecutar `sync_in` (manual o por tarea)
3. Trabajar en codigo
4. Ejecutar `sync_out`
5. Verificar `git status -sb` limpio

## Decisiones operativas
- Fuente de verdad de codigo: GitHub (`main`).
- Sincronizacion GAS: `clasp pull/push` en 2 proyectos:
  - raiz repo
  - `crm-ayudas-subvenciones`
- No forzar merges ni reescribir historial.
- Ante conflicto Git: detener y resolver manualmente.

## Sincronizacion entre dispositivos
- Codigo: GitHub + scripts `sync_*`.
- Hilos Codex: depende de cloud sync de Codex (`codexCloudAccess`).
- Workspace comun obligatorio para reducir desalineaciones.

## Automatizacion instalada
- `CodexSyncIn_5min`: pull continuo cada 5 minutos.
- ONLOGON: si no hay permisos de tarea, fallback Startup:
  - `C:\Users\elrub\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\CodexSyncIn_OnLogon.cmd`

## Riesgos y mitigacion
- Cambios locales no commit + sync_in:
  - mitigacion: bloqueo por defecto; usar `-AllowDirty` solo consciente.
- Diferencias de hilos entre equipos:
  - mitigacion: alinear workspace + completar setup cloud + rediagnosticar.
