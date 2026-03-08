# HILOS-SYNC

## Objetivo
Alinear visibilidad de hilos de Codex entre equipos usando la misma cuenta y el mismo workspace.

## Scripts
- Alinear estado de workspace:
  - `powershell -NoProfile -ExecutionPolicy Bypass -File .\codex-tools\alinear_hilos_codex.ps1`
- Diagnosticar estado real:
  - `powershell -NoProfile -ExecutionPolicy Bypass -File .\codex-tools\diagnostico_hilos_codex.ps1`

## Requisitos en ambos equipos
- Misma cuenta Codex: `manager@rubencoton.com`
- Mismo `account_id`: `64822acc-b670-4d3b-b18f-3fd82e58ee85`
- Mismo workspace activo/guardado:
  - `C:\Users\elrub\Desktop\CARPETA CODEX`

## Cloud de hilos
Si `CLOUD_ACCESS` no es `enabled` (por ejemplo `enabled_needs_setup`):
1. Abrir la app Codex.
2. Ir a ajustes de cuenta/cloud sync.
3. Completar setup cloud con la misma cuenta.
4. Reiniciar app.
5. Re-ejecutar `diagnostico_hilos_codex.ps1`.

## Limitacion real
Sin `codexCloudAccess=enabled` no existe garantia de sincronizacion absoluta de hilos entre dispositivos, aunque el workspace quede alineado.
