# AUDITORIA PROFUNDA - 2026-03-15

## Resultado general
- Estado: OK con mejoras aplicadas
- Auditoria automatica: PASS (58/58)
- Auditoria ultra: PASS (28/28)

## Mejoras aplicadas
1. `CRM_FESTIVALES_ENGINE.gs`
- Ajuste visual rapido en todas las pestanas de festivales (limite de tiempo <= 4.5s).
- Se evita aplicar cambios cuando la cabecera no esta en orden canonico para no romper columnas.
- Refuerzo de validaciones de `GENERO` y `REVISION EMAIL` en modo rapido.
- Limpieza forzada de validacion en `Merge status` para quitar desplegable.
- Soporte de nombres de pestana con alias (`URBANO`, `ELECTRONICA`, `FLAMENCO`).
- Correccion de colision entre `REVISION EMAIL` y `Merge status` en `buildHeaderMap_`.

2. `package.json` + `codex-tools/run-node.ps1`
- `npm run audit:auto` y `npm run audit:ultra` ahora funcionan aunque `node` no este en PATH global.

3. `crm-ayudas-subvenciones`
- Menu con accion manual de limpieza operativa.
- Limpieza onOpen con control por intervalo (6h) y reporte guardado.
- Normalizacion de email mas robusta (placeholders, textos tipo arroba/punto, multiemail).
- Deteccion de pestanas residuales mejorada.

## Verificaciones de sincronizacion
- `gas:push`: OK (subidos `appsscript.json`, `CRM_FESTIVALES_EMAIL_REVIEW.gs`, `CRM_FESTIVALES_ENGINE.gs`).
- `gas:deploy`: BLOQUEADO por dominio (`Only users in the same domain as the script owner may deploy this script`).

## Inspeccion de hoja vinculada
- Intento `clasp run codexInspeccionarHojaVinculada`: BLOQUEADO.
- Motivo: script no desplegado como API executable en el dominio permitido.
- Evidencia disponible en reportes fallback de `codex-tools/reports/`.

## Riesgos abiertos
- Mientras no se habilite deploy API ejecutable, la inspeccion completa remota solo puede hacerse por fallback.
- `Merge status` se protege en codigo, pero si una pestana no esta en estructura canonica no se fuerza en onOpen rapido (se evita riesgo de romper datos).
