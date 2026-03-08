# HILOS-SYNC

## Objetivo
Que Codex muestre los mismos hilos en portatil y sobremesa (en la medida que la nube de Codex lo permita).

## Paso 1 (en ambos equipos)
Ejecutar:
`powershell -ExecutionPolicy Bypass -File .\codex-tools\alinear_hilos_codex.ps1`

## Paso 2 (en ambos equipos)
Ejecutar:
`powershell -ExecutionPolicy Bypass -File .\codex-tools\diagnostico_hilos_codex.ps1`

## Paso 3 (en la app Codex)
- Iniciar sesion con la misma cuenta.
- Abrir exactamente el mismo workspace: `C:\Users\elrub\Desktop\CARPETA CODEX`.
- Completar setup de cloud sync si aparece pendiente.

## Comprobacion rapida
- `CLOUD_ACCESS` deberia dejar de estar en `enabled_needs_setup`.
- `CANONICAL_IN_ACTIVE` y `CANONICAL_IN_SAVED` deben salir `True`.
- El listado de hilos debe verse igual tras reiniciar la app en ambos equipos.

## Limitacion real
Si Codex cloud no termina setup o hay restriccion de plataforma, no existe forma 100% garantizada por script para forzar sincronizacion absoluta de hilos.
