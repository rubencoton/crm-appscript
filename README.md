# CRM FESTIVALES - ARTES BUHO

Proyecto de Google Apps Script para CRM de festivales.
Este README sirve como guia rapida para cualquier persona o IA.

## 1) Contexto
- Empresa: `ARTES BUHO`
- Desarrollador: `RUBEN COTON`
- Objetivo:
  - limpiar datos,
  - auditar emails con contraste web,
  - autocompletar huecos con IA,
  - mantener `Merge status` intacto para YAMM.

## 2) Regla critica
- `Merge status`:
  - siempre ultima columna,
  - sin desplegable,
  - nunca se toca en auditoria/autocompletado.

## 3) Comportamiento al abrir la hoja (`onOpen`)
- No se ejecuta auditoria de contactos automaticamente.
- Si aplica ajustes visuales rapidos en pestanas de festivales.
- Tiempo objetivo: <= 5 segundos.
- Seguridad extra:
  - prioriza la pestana activa para que el cambio se vea al instante,
  - si detecta cabecera incompleta, repara minimo (`REVISION EMAIL` + `Merge status`) sin romper datos,
  - usa cursor rotativo para no quedarse siempre en las mismas pestanas cuando hay limite de tiempo,
  - en pestanas no activas aplica modo ligero (solo cabecera) para recorrer muchas hojas en pocos segundos.

## 4) Menu operativo en Google Sheets
Menu: `CRM FESTIVALES | RUBEN COTON`

1. `BOTON | Auditar contactos web (bloquea + progreso)`
- Auditoria manual de emails con scraping.
- Bloqueo temporal durante ejecucion.
- Progreso 0% a 100%.
- Sobrescribe estado previo de revision.

Estados:
- `BIEN` -> fila verde.
- `CORREGIDO` -> fila azul.
- `MAL` -> fila roja.

2. `BOTON | Autocompletado de celdas (IA)`
- Completa celdas vacias con datos encontrados en web.
- Si no encuentra datos: `IA NO ENCUENTRA`.
- Nunca toca `Merge status`.

## 5) Progreso visible
- Pestana: `PROGRESO_CRM`
- Muestra:
  - estado,
  - porcentaje,
  - barra,
  - ultima actualizacion,
  - bloqueo SI/NO.
- Dialogo de auditoria con estilo futurista.

## 6) Archivos clave
- `CRM_FESTIVALES_ENGINE.gs`
  - menu, onOpen, formato, cabeceras, validaciones.
- `CRM_FESTIVALES_EMAIL_REVIEW.gs`
  - auditoria emails, scraping, bloqueo, progreso, autocompletado.
- `INSPECCION_HOJA.gs`
  - inspeccion tecnica de estructura, formatos, validaciones y protecciones.
- `codex-tools/audit_auto.js`
  - auditoria automatica local.
- `codex-tools/audit_ultra.js`
  - auditoria profunda con stress y validacion live.
- `codex-tools/run-node.ps1`
  - lanzador robusto de Node para Windows cuando `node` no esta en PATH.
- `codex-tools/auditar_sheet_fallback.py`
  - auditoria completa de hoja por XLSX cuando la API ejecutable no esta disponible.

## 7) Cambios importantes de la auditoria profunda (2026-03-15)
- Corregido mapeo de cabeceras para evitar colision entre:
  - `REVISION EMAIL`
  - `Merge status`
- Mejorado `onOpen` para cubrir todas las pestanas de festivales sin romper estructura.
- Ampliado soporte de nombres de pestana:
  - `URBANO`, `ELECTRONICA`, `FLAMENCO` y alias.
- `npm run audit:auto` y `npm run audit:ultra` ahora funcionan sin depender de PATH global.

## 8) Comandos utiles
- `git status -sb`
- `npm run gas:status`
- `npm run gas:push`
- `npm run gas:deploy`
- `npm run audit:auto`
- `npm run audit:ultra`
- `python codex-tools\\auditar_sheet_fallback.py --spreadsheet-id <ID> --xlsx-path <RUTA_XLSX>`
- `powershell -NoProfile -ExecutionPolicy Bypass -File .\\codex-tools\\auditoria_festivales_2h.ps1 -Horas 2 -IntervaloMin 10`

## 9) Trazabilidad
- Reportes en `codex-tools/reports/`.
- Regla operativa:
  - actualizar README,
  - commit en GitHub,
  - push Apps Script,
  - registrar resultado (ok/error).

## 10) Limitaciones conocidas
- `gas:deploy` puede fallar por dominio:
  - `Only users in the same domain as the script owner may deploy this script.`
- `INSPECCION_HOJA.gs` via `clasp run` requiere despliegue API ejecutable.
- Si no hay API ejecutable disponible, usar reporte fallback en `codex-tools/reports/`.

## 11) Script ID
- Apps Script ID: `1OGuPezQ26BFvaLRiy-IYIotGpmVu_Z_b9Mi8tCiprIz8zB4DgqmMc5Ea`
- Vinculo local: `.clasp.json`

## 12) Mantenimiento
Siempre:
- actualizar README,
- versionar en GitHub,
- sincronizar Apps Script,
- documentar bloqueos de permisos si aparecen.
