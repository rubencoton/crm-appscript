# AGENTS.md

## Regla persistente para proyectos de Google Apps Script (vinculados a Google Sheets)

Siempre que el usuario pida analizar, revisar o modificar un proyecto de Apps Script vinculado a una hoja de calculo, antes de proponer cambios o conclusiones se debe ejecutar una inspeccion completa de la hoja asociada y reportar:

1. Estructura completa (libro, pestanas, rangos usados, filas/columnas).
2. Datos visibles relevantes y formulas.
3. Formatos aplicados (colores, tipografias, negrita/cursiva, bordes, alineacion, formatos de numero/fecha/moneda, tamanos).
4. Celdas combinadas.
5. Validaciones de datos y desplegables.
6. Filtros y vistas de filtro.
7. Reglas de formato condicional.
8. Protecciones de hoja/rango y permisos aplicados.

Si no hay acceso directo a la hoja, solicitar o ejecutar el mecanismo tecnico necesario (Apps Script de inspeccion, exportacion o evidencias) antes de cerrar el analisis.
