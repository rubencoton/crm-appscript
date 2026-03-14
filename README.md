# CRM FESTIVALES - ARTES BUHO

Proyecto de Google Apps Script para el CRM de festivales.
Este README esta pensado para que cualquier persona (o IA) entienda el sistema al abrirlo desde cualquier hilo o equipo.

## 1) Contexto del proyecto
- Empresa: `ARTES BUHO`
- Desarrollador: `RUBEN COTON`
- Objetivo principal:
  - mantener datos limpios,
  - auditar emails con contraste web,
  - autocompletar vacios con ayuda IA,
  - preservar columna de merge para YAMM.

## 2) Regla critica de negocio
- La columna `Merge status`:
  - siempre es la ultima columna,
  - no se modifica en auditoria ni autocompletado,
  - no usa desplegable.

## 3) Comportamiento al abrir la hoja
- `onOpen` NO ejecuta auditoria de contactos.
- `onOpen` solo aplica ajustes visuales ligeros del sheet activo.
- Objetivo de apertura: respuesta rapida (meta <= 5 segundos).

## 4) Menu operativo (arriba en Google Sheets)
Menu: `CRM FESTIVALES | RUBEN COTON`

1. `BOTON | Auditar contactos web (bloquea + progreso)`
- Ejecuta auditoria manual de emails con scraping.
- Bloquea temporalmente la hoja mientras corre.
- Muestra progreso de 0% a 100%.
- Sobrescribe estado previo de revision en cada fila.

Estados y color por fila:
- `BIEN` -> verde (email validado).
- `CORREGIDO` -> azul (se encontro y sustituyo email alternativo).
- `MAL` -> rojo (no validable o no recuperable).

2. `BOTON | Autocompletado de celdas (IA)`
- Recorre celdas vacias y busca datos en web.
- Si encuentra dato fiable, lo escribe.
- Si no encuentra dato, escribe `IA NO ENCUENTRA`.
- Nunca toca `Merge status`.

## 5) Progreso en tiempo real
- Pestaña de progreso: `PROGRESO_CRM`.
- Datos visibles:
  - estado actual,
  - porcentaje,
  - barra de progreso,
  - ultima actualizacion,
  - bloqueo activo SI/NO.
- Tambien hay panel emergente con estilo futurista para la auditoria manual.

## 6) Archivos clave y para que sirve cada uno
- `CRM_FESTIVALES_ENGINE.gs`
  - menu, onOpen, formato visual, cabeceras, validaciones, reglas de estructura.
- `CRM_FESTIVALES_EMAIL_REVIEW.gs`
  - auditoria de emails, scraping, bloqueo, progreso, autocompletado IA.
- `INSPECCION_HOJA.gs`
  - inspeccion tecnica de estructura/formato/validaciones/protecciones.
- `codex-tools/audit_auto.js`
  - auditoria tecnica automatica (sintaxis, secretos, stress local).
- `codex-tools/audit_ultra.js`
  - auditoria profunda con checks avanzados.
- `codex-tools/auditoria_festivales_2h.ps1`
  - ciclo de auditoria continua 2h con logs y csv de trazabilidad.

## 7) Flujo recomendado de trabajo
1. Cambiar codigo en repo.
2. Actualizar README cuando cambie funcionalidad.
3. Commit con mensaje claro.
4. `git push`.
5. `gas:push` para subir a Apps Script.
6. Verificar en hoja real usando botones del menu.

## 8) Comandos utiles
- Estado Git:
  - `git status -sb`
- Estado Apps Script:
  - `npm run gas:status`
- Subir codigo a Apps Script:
  - `npm run gas:push`
- Deploy (si el dominio lo permite):
  - `npm run gas:deploy`
- Auditoria tecnica:
  - `node codex-tools/audit_auto.js`
  - `node codex-tools/audit_ultra.js`
- Auditoria continua 2h:
  - `powershell -NoProfile -ExecutionPolicy Bypass -File .\codex-tools\auditoria_festivales_2h.ps1 -Horas 2 -IntervaloMin 10`

## 9) Trazabilidad y reportes
- Reportes en `codex-tools/reports/`.
- Cada cambio funcional debe dejar:
  - commit en GitHub,
  - README actualizado,
  - nota de estado de push a Apps Script.

## 10) Limitaciones conocidas
- `gas:deploy` puede fallar por restriccion de dominio:
  - mensaje tipico: `Only users in the same domain as the script owner may deploy this script.`
- `clasp push` puede requerir reautenticacion Google (error `invalid_rapt`).

## 11) Script ID y enlace tecnico
- Script ID (Apps Script): `1OGuPezQ26BFvaLRiy-IYIotGpmVu_Z_b9Mi8tCiprIz8zB4DgqmMc5Ea`
- Vinculacion local: `.clasp.json`

## 12) Regla de mantenimiento
Siempre que se toque funcionalidad:
- se actualiza este README,
- se deja commit en GitHub,
- se intenta push a Apps Script y se reporta resultado.
