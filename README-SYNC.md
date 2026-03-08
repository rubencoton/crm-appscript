# README-SYNC

## 1) Traer cambios (antes de editar)
- Doble clic en `sync_in.bat`

## 2) Subir cambios (despues de editar)
- Doble clic en `sync_out.bat`

## 3) Hacer todo junto
- Doble clic en `sync_full.bat`

## 4) Desde terminal
- Entrada: `powershell -ExecutionPolicy Bypass -File .\sync_in.ps1`
- Salida: `powershell -ExecutionPolicy Bypass -File .\sync_out.ps1`
- Completo: `powershell -ExecutionPolicy Bypass -File .\sync_full.ps1`

## 5) Regla de seguridad
- Si Git muestra conflicto/rebase bloqueado, no forzar.
- Resolver manualmente o abortar con `git rebase --abort`.

## 6) Automatizacion
- Tarea cada 5 minutos: `CodexSyncInFestivales_Every5Min`
- Inicio de sesion (fallback): `C:\Users\elrub\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\CodexSyncInOnLogon.bat`
