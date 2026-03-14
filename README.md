# CRM FESTIVALES - ARTES BUHO

Aplicacion de Google Apps Script conectada a Google Sheets para gestionar y auditar contactos de festivales.

## Autor y empresa
- Desarrollador: `RUBEN COTON`
- Empresa: `ARTES BUHO`

## Objetivo del sistema
- Mantener la hoja CRM homogénea y util para operativa comercial.
- Auditar emails de contactos con contraste web.
- Reducir errores en datos clave sin tocar la columna de merge para YAMM.

## Regla clave de columnas
- La columna `Merge status` debe ser siempre la ultima.
- El contenido de `Merge status` no se modifica durante auditorias ni autocompletado.
- `Merge status` no usa desplegable (campo libre).

## Comportamiento al abrir la hoja (onOpen)
- Se aplican solo ajustes visuales rapidos del sheet activo.
- Objetivo de tiempo: no superar ~5 segundos.
- No se ejecuta auditoria de contactos automaticamente.
- Se limpian triggers de auto-auditoria para forzar modo manual por boton.

## Botones del menu CRM
Menu: `CRM FESTIVALES | RUBEN COTON`

1. `BOTON | Auditar contactos web (bloquea + progreso)`
- Lanza auditoria manual de emails fila a fila.
- Bloquea temporalmente la hoja mientras corre.
- Muestra progreso 0%-100% con panel emergente.
- Estados finales por fila:
  - `BIEN` -> fila verde.
  - `CORREGIDO` -> fila azul (se encontro email alternativo y se sustituyo).
  - `MAL` -> fila roja (no valido / no recuperable).
- La auditoria sobrescribe el estado previo de revision (no reutiliza lo que habia).

2. `BOTON | Autocompletado de celdas (IA)`
- Recorre celdas vacias y busca datos por scraping web.
- Si encuentra dato valido, lo escribe.
- Si no encuentra nada, escribe `IA NO ENCUENTRA`.
- No modifica `Merge status`.

## Pestaña de progreso
- Nombre: `PROGRESO_CRM`
- Muestra:
  - estado actual,
  - porcentaje,
  - barra visual,
  - timestamp,
  - indicador de bloqueo.

## Ficheros principales
- `CRM_FESTIVALES_ENGINE.gs`
  - menu, onOpen, formato visual, normalizacion base.
- `CRM_FESTIVALES_EMAIL_REVIEW.gs`
  - auditoria de contactos/email, progreso, bloqueo, autocompletado IA.
- `INSPECCION_HOJA.gs`
  - utilidades de inspeccion completa (estructura/formato/validaciones/etc).

## Sincronizacion
- ScriptId Apps Script: `1OGuPezQ26BFvaLRiy-IYIotGpmVu_Z_b9Mi8tCiprIz8zB4DgqmMc5Ea`
- Comandos:
  - `npm run gas:status`
  - `npm run gas:pull`
  - `npm run gas:push`
  - `npm run gas:deploy`

## Nota de despliegue
- Si `gas:deploy` devuelve error de dominio (`Only users in the same domain...`), el codigo puede estar igualmente subido con `gas:push`.
- En ese caso, validar cambios abriendo la hoja vinculada y usando los botones del menu.
