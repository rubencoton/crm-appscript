# README-SYNC

## Uso diario (muy simple)
1. Doble clic en `sync_in.bat` para traer cambios de GitHub y Apps Script.
2. Edita tus archivos.
3. Doble clic en `sync_out.bat` para subir a Apps Script y GitHub.
4. Si quieres hacerlo de una vez, usa `sync_full.bat`.

## Regla de oro
Antes de editar en cualquier dispositivo, ejecuta siempre `sync_in` primero.

## Conflictos Git
Si aparece conflicto, no fuerces nada. Resuelve manualmente o aborta rebase con `git rebase --abort`.
