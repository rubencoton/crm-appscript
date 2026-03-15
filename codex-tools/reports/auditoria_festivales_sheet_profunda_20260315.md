# AUDITORIA PROFUNDA FESTIVALES (HOJA + CODIGO)

- Fecha: 2026-03-15
- Hoja auditada: 1kRrdCwd0n6FwVp-8rKP3gEeYT-TBXc1JL8sa_xzp1IM
- Metodo hoja: fallback XLSX (sin API executable disponible)

## Resultado tecnico
- audit:auto: PASS 58/58
- audit:ultra: PASS 28/28

## Hallazgos hoja (snapshot)
- Total pestanas: 32
- Pestanas con filtro activo: 31
- Pestanas con errores de formula visibles: 0
- Pestanas sin validaciones detectables: 23
- Cabecera sin `REVISION EMAIL`: 23 pestanas (principalmente ROCK/ELECTR/JAZZ/FLAM/RUMBA/MR/MC/PTE)

## Mejoras aplicadas en codigo
1. `CRM_FESTIVALES_ENGINE.gs`
- onOpen ahora prioriza la pestana activa.
- onOpen intenta reparacion minima de cabecera cuando faltan columnas clave.
- Reparacion minima automatica:
  - crea `REVISION EMAIL` si falta,
  - mantiene/asegura `Merge status`,
  - quita validacion desplegable de `Merge status`.

2. `crm-ayudas-subvenciones/.claspignore`
- Corregido include esperado por auditor (`!DecisionInstantanea.gs`).

3. `codex-tools/auditar_sheet_fallback.py`
- Nuevo script para inspeccion completa de hoja por XLSX.
- Cubre: estructura, datos/formulas, formatos, combinadas, validaciones, filtros, condicionales y protecciones (nivel sheet).

## Limitaciones abiertas
- `clasp run codexInspeccionarHojaVinculada` sigue bloqueado por: API executable no disponible en dominio.
- `gas:deploy` sigue bloqueado por dominio del owner.

## Siguiente ejecucion recomendada
- Abrir la hoja y dejar que onOpen repare cabeceras minimas por tandas.
- Ejecutar despues `BOTON | Auditar contactos web (bloquea + progreso)` para coloreado y estado BIEN/CORREGIDO/MAL.
