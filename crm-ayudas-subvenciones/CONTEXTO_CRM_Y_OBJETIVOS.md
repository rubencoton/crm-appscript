# CONTEXTO CRM AYUDAS Y SUBVENCIONES

Este documento resume que se ha pedido, por que se ha pedido y como esta implementado para que cualquier chat futuro pueda continuar sin perder contexto.

## Objetivo de negocio

- Tener un CRM de ayudas/subvenciones que permita decidir rapido, aunque falten datos.
- Que una persona no tecnica pueda entender la situacion de un vistazo.
- Evitar procesos lentos: la hoja debe reaccionar al instante al editar.

## Reglas clave

1. La `FECHA LIMITE` es el dato mas importante.
2. Si no hay fecha oficial clara, se estima y se marca como `ESTIMADO: DD/MM/YYYY`.
3. Inscripcion `ABIERTA` cuando hoy esta entre `(fecha limite - 3 meses)` y `fecha limite`.
4. Inscripcion `CERRADA` cuando la fecha limite ya paso.
5. `SIN PUBLICAR` cuando no hay evidencia suficiente.

## Reglas visuales

- Fila en verde claro cuando esta `ABIERTA`.
- Fila en rojo claro cuando esta `CERRADA`.
- Fila en amarillo cuando esta `SIN PUBLICAR`.
- Campos criticos vacios se resaltan para revisar calidad de dato.

## IA (modelo y comportamiento)

- Modelo prioritario: `gemini-3.1-pro-preview`.
- El prompt exige razonamiento profundo:
  - detectar contradicciones,
  - explicar supuestos,
  - informar nivel de confianza.
- La salida se normaliza para que siempre tenga estructura util.

## Estabilidad tecnica aplicada

- Reintentos con backoff para llamadas IA.
- Parseo JSON defensivo para no romper ejecuciones.
- Lock en `onEdit` para evitar condiciones de carrera.
- Eliminacion de archivos duplicados para evitar colisiones de funciones.

## Archivos principales

- `Code.js`: motor principal (menu, escaneo, IA, formato, reglas).
- `DecisionInstantanea.gs`: capa de decision instantanea y escenarios.
- `appsscript.json`: scopes y runtime Apps Script.

## Criterio de continuidad en futuros chats

Si se abre otro chat, partir de estas prioridades:

1. No romper la regla de fecha limite + ventana de 3 meses.
2. Mantener respuesta instantanea al editar.
3. Mantener un solo motor principal (evitar duplicados funcionales).
4. Documentar cualquier cambio de negocio en este archivo.
