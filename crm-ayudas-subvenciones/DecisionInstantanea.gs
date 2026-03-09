// =============================================================================
// CRM AYUDAS - CAPA DE DECISION INSTANTANEA
// =============================================================================

const CRM_MODE_PROP = 'INSTANT_MODE';
const CRM_LOCK_PREFIX = 'CRM_LOCKED_AUTOFILL';

const CRM_DECISION = {
  SHEET_ESCENARIOS: 'ESCENARIOS',
  SHEET_PANEL: 'PANEL_DECISION',
  SCENARIO_HEADERS: ['ESCENARIO', 'FACTOR_VOLUMEN', 'FACTOR_EXITO', 'FACTOR_CARGA', 'FOCO'],
  SCENARIO_DEFAULTS: [
    ['CONSERVADOR', 0.7, 0.8, 0.7, 'Asegurar calidad y reducir riesgo.'],
    ['BASE', 1.0, 1.0, 1.0, 'Equilibrio entre crecimiento y control.'],
    ['AMBICIOSO', 1.35, 1.2, 1.25, 'Maximizar oportunidades y visibilidad.']
  ],
  LOCKED_COLUMNS: [CRM_COL.ESTADO, CRM_COL.INSCRIPCION, CRM_COL.FECHA_DESARROLLO]
};

function prepararHojaDecisionInteligente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  prepararEntornoDecision_(ss, true);
}

function actualizarPanelDecision() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  recalcularTablasDecision_(ss, true);
  SpreadsheetApp.getUi().alert('Panel de escenarios actualizado.\n' + FIRMA_APP);
}

function aplicarBloqueoCeldasCalculadas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CRM_CONFIG.SHEET_CONCURSOS);
  if (!sh) throw new Error('No existe la pestana CONCURSOS.');
  asegurarBloqueosCalculados_(sh, true);
  SpreadsheetApp.getUi().alert('Celdas calculadas bloqueadas.\n' + FIRMA_APP);
}

function quitarBloqueoCeldasCalculadas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  quitarBloqueosCalculados_(ss.getSheetByName(CRM_CONFIG.SHEET_CONCURSOS));
  quitarBloqueoPanelDecision_(ss.getSheetByName(CRM_DECISION.SHEET_PANEL));
  SpreadsheetApp.getUi().alert('Bloqueos retirados en CONCURSOS/PANEL_DECISION.\n' + FIRMA_APP);
}

function activarModoInstantaneo() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(CRM_MODE_PROP, 'TRUE');
  if (typeof limpiarTriggersEjecucion_ === 'function') {
    limpiarTriggersEjecucion_();
  }
  SpreadsheetApp.getUi().alert('Modo instantaneo activado.\nLas acciones arrancan al momento y sin espera por temporizador.\n' + FIRMA_APP);
}

function activarModoAsincrono() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(CRM_MODE_PROP, 'FALSE');
  SpreadsheetApp.getUi().alert('Modo asincrono activado.\nLas acciones volveran a usar temporizador de arranque.\n' + FIRMA_APP);
}

function usarModoInstantaneoCRM_() {
  return sanitizeValue_(PropertiesService.getScriptProperties().getProperty(CRM_MODE_PROP)) === 'TRUE';
}

function onEscaneoFinalizadoDecision_(ss) {
  try {
    recalcularTablasDecision_(ss || SpreadsheetApp.getActiveSpreadsheet(), false);
  } catch (err) {
    logCRM_('No se pudo refrescar PANEL_DECISION al finalizar: ' + err.message, 'warning');
  }
}

function actualizarEdicionInstantaneaCRM_(sheet, rowN, editedCol, ss) {
  if (!sheet || sheet.getName() !== CRM_CONFIG.SHEET_CONCURSOS || rowN < 2) return false;
  const estado = autocompletarFilaDecision_(sheet, rowN, editedCol);
  const tz = (ss || SpreadsheetApp.getActive()).getSpreadsheetTimeZone();
  const ins = sanitizeValue_(sheet.getRange(rowN, CRM_COL.INSCRIPCION).getValue()).toUpperCase();

  aplicarFormatoFila_(sheet, rowN, ins, tz);
  aplicarColorEstadoDecision_(sheet.getRange(rowN, CRM_COL.ESTADO), estado || sheet.getRange(rowN, CRM_COL.ESTADO).getValue());
  asegurarBloqueosCalculados_(sheet, false);

  try {
    recalcularTablasDecision_(sheet.getParent(), false);
  } catch (err) {
    logCRM_('No se pudo recalcular panel tras edicion: ' + err.message, 'warning');
  }

  return true;
}

function autocompletarFilaDecision_(sheet, rowN, editedCol) {
  const rowRange = sheet.getRange(rowN, 1, 1, 17);
  const row = rowRange.getValues()[0];
  if (filaEstaVaciaDecision_(row)) return '';

  let changed = false;
  const nombre = sanitizeValue_(row[CRM_COL.NOMBRE - 1]);

  const estadoRaw = sanitizeValue_(row[CRM_COL.ESTADO - 1]);
  let estado = normalizeEstado_(estadoRaw);
  if (!estadoRaw && nombre) estado = CRM_ESTADO.REVISAR;
  if (estadoRaw !== estado) {
    row[CRM_COL.ESTADO - 1] = estado;
    changed = true;
  }

  const fechaOriginal = displayCell_(row[CRM_COL.FECHA_LIMITE - 1]);
  const fechaNormalizada = normalizarFechaLimite_(fechaOriginal, fechaOriginal);
  if (sanitizeValue_(fechaOriginal) !== sanitizeValue_(fechaNormalizada)) {
    row[CRM_COL.FECHA_LIMITE - 1] = fechaNormalizada;
    changed = true;
  } else if (!sanitizeValue_(row[CRM_COL.FECHA_LIMITE - 1])) {
    row[CRM_COL.FECHA_LIMITE - 1] = fechaNormalizada;
    changed = true;
  }

  const inscripcion = estadoInscripcionDesdeFecha_(fechaNormalizada);
  const insRaw = sanitizeValue_(row[CRM_COL.INSCRIPCION - 1]).toUpperCase();
  if (insRaw !== inscripcion) {
    row[CRM_COL.INSCRIPCION - 1] = inscripcion;
    changed = true;
  }

  const desarrolloActual = sanitizeValue_(row[CRM_COL.FECHA_DESARROLLO - 1]);
  const desarrolloSugerido = normalizarMesDesarrollo_(desarrolloActual || fechaNormalizada);
  if (!desarrolloActual || desarrolloActual !== desarrolloSugerido || editedCol === CRM_COL.FECHA_LIMITE) {
    row[CRM_COL.FECHA_DESARROLLO - 1] = desarrolloSugerido;
    changed = true;
  }

  const tipoRaw = sanitizeValue_(row[CRM_COL.TIPO_PREMIO - 1]);
  const tipoNorm = normalizarTipoPremio_(tipoRaw || 'VARIOS');
  if (tipoRaw !== tipoNorm) {
    row[CRM_COL.TIPO_PREMIO - 1] = tipoNorm;
    changed = true;
  }

  const paisRaw = sanitizeValue_(row[CRM_COL.PAIS - 1]);
  if (!paisRaw && (sanitizeValue_(row[CRM_COL.MUNICIPIO - 1]) || sanitizeValue_(row[CRM_COL.PROVINCIA - 1]) || nombre)) {
    row[CRM_COL.PAIS - 1] = 'España';
    changed = true;
  }

  for (let c = CRM_COL.LINK1; c <= CRM_COL.LINK3; c++) {
    const idx = c - 1;
    const raw = sanitizeValue_(row[idx]);
    if (!raw) continue;
    const url = normalizarUrl_(raw);
    if (url && url !== raw) {
      row[idx] = url;
      changed = true;
    }
  }

  if (changed) {
    rowRange.setValues([row]);
  }

  return row[CRM_COL.ESTADO - 1];
}

function aplicarColorEstadoDecision_(cell, estadoRaw) {
  const estado = normalizeEstado_(estadoRaw);
  let bg = '#EFEFEF';
  let fc = '#333333';
  if (estado === CRM_ESTADO.REVISADO_IA) { bg = '#D9EAD3'; fc = '#1E4620'; }
  if (estado === CRM_ESTADO.REVISADO_HUMANO) { bg = '#D0E0F5'; fc = '#113A67'; }
  if (estado === CRM_ESTADO.NUEVO_DESCUBRIMIENTO) { bg = '#FFE9B8'; fc = '#704D00'; }
  cell.setBackground(bg).setFontColor(fc);
}

function filaEstaVaciaDecision_(row) {
  for (let i = 0; i < row.length; i++) {
    if (sanitizeValue_(row[i])) return false;
  }
  return true;
}

function prepararEntornoDecision_(ss, showAlert) {
  if (!ss) return;
  const shConcursos = ss.getSheetByName(CRM_CONFIG.SHEET_CONCURSOS);
  if (!shConcursos) {
    throw new Error('No existe la pestana CONCURSOS.');
  }

  asegurarCabeceraConcursosDecision_(shConcursos);
  const shEscenarios = asegurarHojaEscenarios_(ss);
  const shPanel = asegurarHojaPanelDecision_(ss);
  asegurarBloqueosCalculados_(shConcursos, true);
  bloquearPanelDecision_(shPanel);
  recalcularTablasDecision_(ss, true);

  if (showAlert) {
    SpreadsheetApp.getUi().alert(
      'Hoja preparada para decision en tiempo real.\n\n' +
      '1) Edita CONCURSOS y el sistema autocompleta al instante.\n' +
      '2) Ajusta factores en ESCENARIOS.\n' +
      '3) Revisa PANEL_DECISION para elegir estrategia.'
    );
  }

  return { escenarios: shEscenarios.getName(), panel: shPanel.getName() };
}

function asegurarCabeceraConcursosDecision_(sheet) {
  const expected = [
    'NOMBRE CONCURSO', 'ESTADO', 'INSCRIPCION', 'FECHA LIMITE', 'FECHA DESARROLLO',
    'TIPO PREMIO', 'DETALLE PREMIO', 'DIRIGIDO A', 'MUNICIPIO', 'PROVINCIA',
    'PAIS', 'LINK1', 'LINK2', 'LINK3', 'EMAIL', 'TELEFONO', 'NOTAS'
  ];

  const headerRange = sheet.getRange(1, 1, 1, expected.length);
  const current = headerRange.getDisplayValues()[0];
  let emptyCount = 0;
  for (let i = 0; i < current.length; i++) {
    if (!sanitizeValue_(current[i])) emptyCount++;
  }

  if (emptyCount >= 8) {
    headerRange.setValues([expected]);
  }

  headerRange
    .setBackground('#6B0018')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
}

function asegurarHojaEscenarios_(ss) {
  let sh = ss.getSheetByName(CRM_DECISION.SHEET_ESCENARIOS);
  if (!sh) sh = ss.insertSheet(CRM_DECISION.SHEET_ESCENARIOS);

  if (sh.getMaxColumns() < 8) {
    sh.insertColumnsAfter(sh.getMaxColumns(), 8 - sh.getMaxColumns());
  }

  sh.getRange(1, 1, 1, 5).setValues([CRM_DECISION.SCENARIO_HEADERS]);

  const existing = sh.getRange(2, 1, 3, 5).getValues();
  const rows = [];
  for (let i = 0; i < CRM_DECISION.SCENARIO_DEFAULTS.length; i++) {
    const def = CRM_DECISION.SCENARIO_DEFAULTS[i];
    const ex = existing[i] || [];

    const name = sanitizeValue_(ex[0]) || def[0];
    const fVol = Number(ex[1]);
    const fExito = Number(ex[2]);
    const fCarga = Number(ex[3]);
    const foco = sanitizeValue_(ex[4]) || def[4];

    rows.push([
      name,
      isFinite(fVol) && fVol > 0 ? fVol : def[1],
      isFinite(fExito) && fExito > 0 ? fExito : def[2],
      isFinite(fCarga) && fCarga > 0 ? fCarga : def[3],
      foco
    ]);
  }

  sh.getRange(2, 1, rows.length, 5).setValues(rows);

  sh.getRange(1, 1, 1, 5)
    .setBackground('#1F2937')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sh.getRange(2, 2, 3, 3).setNumberFormat('0.00');
  sh.getRange(2, 1, 3, 5)
    .setBackground('#F8FAFC')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID);

  const rule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0.4, 2.5)
    .setAllowInvalid(false)
    .setHelpText('Rango recomendado: 0.40 a 2.50')
    .build();
  sh.getRange(2, 2, 3, 3).setDataValidation(rule);

  sh.getRange('G1').setValue('COMO USAR').setFontWeight('bold').setFontColor('#111827');
  sh.getRange('G2').setValue('1) Ajusta factores en columnas B-D.');
  sh.getRange('G3').setValue('2) El panel se recalcula al instante.');
  sh.getRange('G4').setValue('3) Elige el escenario con mejor balance.');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 7);

  return sh;
}

function asegurarHojaPanelDecision_(ss) {
  let sh = ss.getSheetByName(CRM_DECISION.SHEET_PANEL);
  if (!sh) sh = ss.insertSheet(CRM_DECISION.SHEET_PANEL);
  return sh;
}

function recalcularTablasDecision_(ss, ensureSheets) {
  if (!ss) return null;
  const shConcursos = ss.getSheetByName(CRM_CONFIG.SHEET_CONCURSOS);
  if (!shConcursos) return null;

  let shEscenarios = ss.getSheetByName(CRM_DECISION.SHEET_ESCENARIOS);
  let shPanel = ss.getSheetByName(CRM_DECISION.SHEET_PANEL);

  if (ensureSheets) {
    shEscenarios = asegurarHojaEscenarios_(ss);
    shPanel = asegurarHojaPanelDecision_(ss);
  }

  if (!shEscenarios || !shPanel) return null;

  const metricas = calcularMetricasConcursosDecision_(shConcursos);
  const escenarios = shEscenarios.getRange(2, 1, 3, 5).getValues();
  renderPanelDecision_(shPanel, metricas, escenarios, ss.getSpreadsheetTimeZone());
  bloquearPanelDecision_(shPanel);

  return { metricas: metricas, escenarios: escenarios };
}

function calcularMetricasConcursosDecision_(sheetConcursos) {
  const out = {
    total: 0,
    abiertas: 0,
    cerradas: 0,
    sinPublicar: 0,
    revisadoHumano: 0,
    revisadoIA: 0,
    nuevos: 0,
    contactoOK: 0,
    linksOK: 0,
    proximas30: 0,
    vencidas: 0,
    sinUbicacion: 0,
    scoreCalidad: 0,
    riesgo: 10,
    recomendacionGlobal: 'Primero completa datos base para activar decisiones.'
  };

  const lastRow = sheetConcursos.getLastRow();
  if (lastRow <= 1) return out;

  const data = sheetConcursos.getRange(2, 1, lastRow - 1, 17).getValues();
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const nombre = sanitizeValue_(row[CRM_COL.NOMBRE - 1]);
    if (!nombre) continue;

    out.total++;

    const estado = normalizeEstado_(row[CRM_COL.ESTADO - 1]);
    const ins = sanitizeValue_(row[CRM_COL.INSCRIPCION - 1]).toUpperCase();

    if (estado === CRM_ESTADO.REVISADO_HUMANO) out.revisadoHumano++;
    if (estado === CRM_ESTADO.REVISADO_IA) out.revisadoIA++;
    if (estado === CRM_ESTADO.NUEVO_DESCUBRIMIENTO) out.nuevos++;

    if (ins === CRM_INSCRIPCION.ABIERTA) out.abiertas++;
    else if (ins === CRM_INSCRIPCION.CERRADA) out.cerradas++;
    else out.sinPublicar++;

    const email = sanitizeValue_(row[CRM_COL.EMAIL - 1]);
    const tel = sanitizeValue_(row[CRM_COL.TELEFONO - 1]);
    if (isValidEmail_(email) || !!formatSpanishPhone_(tel)) out.contactoOK++;

    if (
      isValidHttpUrl_(row[CRM_COL.LINK1 - 1]) ||
      isValidHttpUrl_(row[CRM_COL.LINK2 - 1]) ||
      isValidHttpUrl_(row[CRM_COL.LINK3 - 1])
    ) {
      out.linksOK++;
    }

    const municipio = sanitizeValue_(row[CRM_COL.MUNICIPIO - 1]);
    const provincia = sanitizeValue_(row[CRM_COL.PROVINCIA - 1]);
    if (!municipio && !provincia) out.sinUbicacion++;

    const parsed = parseFechaLimite_(row[CRM_COL.FECHA_LIMITE - 1]);
    if (parsed && parsed.date) {
      const limite = new Date(parsed.date.getFullYear(), parsed.date.getMonth(), parsed.date.getDate(), 0, 0, 0, 0);
      const diff = Math.floor((limite.getTime() - hoy.getTime()) / (24 * 60 * 60 * 1000));
      if (diff >= 0 && diff <= 30) out.proximas30++;
      if (diff < 0) out.vencidas++;
    }
  }

  if (!out.total) return out;

  const calidadRaw = ((out.contactoOK + out.linksOK + out.revisadoHumano) / (out.total * 3)) * 100;
  out.scoreCalidad = Math.max(0, Math.min(100, Math.round(calidadRaw)));

  const ratioSinPublicar = out.sinPublicar / out.total;
  const ratioSinUbicacion = out.sinUbicacion / out.total;
  const riesgoRaw = 10 - (out.scoreCalidad / 15) + (ratioSinPublicar * 4) + (ratioSinUbicacion * 2);
  out.riesgo = Math.max(1, Math.min(10, Math.round(riesgoRaw)));

  out.recomendacionGlobal = construirRecomendacionGlobalDecision_(out);
  return out;
}

function construirRecomendacionGlobalDecision_(metricas) {
  if (!metricas.total) return 'Carga inicial: introduce convocatorias para activar el panel.';
  if (metricas.abiertas === 0) return 'No hay convocatorias ABIERTA: prioriza radar y nuevas fuentes.';
  if (metricas.riesgo >= 8) return 'Riesgo alto: limpia datos clave antes de escalar envios.';
  if (metricas.proximas30 >= 6) return 'Ventana intensa: prioriza convocatorias que cierran en <=30 dias.';
  if (metricas.scoreCalidad < 55) return 'Mejorar calidad de contacto y links antes de tomar decisiones agresivas.';
  return 'Situacion saludable: puedes operar con escenario BASE o AMBICIOSO.';
}

function renderPanelDecision_(sheet, metricas, escenarios, tz) {
  sheet.clear();
  sheet.getRange(1, 1, 1, 8).merge();
  sheet.getRange(1, 1).setValue('PANEL DE DECISION | CRM AYUDAS Y SUBVENCIONES');
  sheet.getRange(2, 1, 1, 8).merge();
  sheet.getRange(2, 1).setValue(
    'Actualizado: ' + Utilities.formatDate(new Date(), tz || Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss') +
    ' | Calidad: ' + metricas.scoreCalidad + '/100 | Riesgo: ' + metricas.riesgo + '/10'
  );

  sheet.getRange(1, 1, 1, 8)
    .setBackground('#0F172A')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(13);
  sheet.getRange(2, 1, 1, 8)
    .setBackground('#E2E8F0')
    .setFontColor('#0F172A')
    .setHorizontalAlignment('left')
    .setFontSize(10);

  const resumen = [
    ['INDICADOR', 'VALOR', 'INTERPRETACION'],
    ['Convocatorias totales', metricas.total, 'Universo actual de oportunidades registradas.'],
    ['Convocatorias ABIERTA', metricas.abiertas, 'Pipeline activo para accion inmediata.'],
    ['Cierre en <=30 dias', metricas.proximas30, 'Prioridad operativa de corto plazo.'],
    ['Calidad de datos', metricas.scoreCalidad + '/100', 'Mayor valor = decisiones mas fiables.'],
    ['Riesgo operativo', metricas.riesgo + '/10', 'Mayor valor = mas incertidumbre.'],
    ['Recomendacion global', metricas.recomendacionGlobal, 'Guia para decidir el siguiente movimiento.']
  ];

  sheet.getRange(4, 1, resumen.length, 3).setValues(resumen);
  sheet.getRange(4, 1, 1, 3)
    .setBackground('#1E3A8A')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange(5, 1, resumen.length - 1, 3)
    .setBackground('#F8FAFC')
    .setFontColor('#111827')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID);

  const headerEsc = [['ESCENARIO', 'OBJETIVO ABIERTAS', 'PROB. EXITO', 'OPORT. ESPERADAS', 'CARGA (1-10)', 'RIESGO (1-10)', 'DECISION', 'PRIMER PASO']];
  sheet.getRange(13, 1, 1, 8).setValues(headerEsc);
  sheet.getRange(13, 1, 1, 8)
    .setBackground('#7C2D12')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const baseRate = Math.max(0.18, Math.min(0.85, (metricas.scoreCalidad / 100) * 0.75));
  const rowsEsc = [];

  for (let i = 0; i < escenarios.length; i++) {
    const r = escenarios[i] || [];
    const name = sanitizeValue_(r[0]) || ('ESCENARIO ' + (i + 1));
    const fVol = Number(r[1]) || 1;
    const fExito = Number(r[2]) || 1;
    const fCarga = Number(r[3]) || 1;

    const objetivo = Math.max(0, Math.round(metricas.abiertas * fVol));
    const prob = Math.max(0.05, Math.min(0.95, baseRate * fExito));
    const oportunidades = Math.round(objetivo * prob);
    const carga = Math.max(1, Math.min(10, Math.round(((objetivo / Math.max(1, metricas.abiertas || 1)) * 5 * fCarga) + 2)));
    const riesgo = Math.max(1, Math.min(10, Math.round(metricas.riesgo + ((fVol - 1) * 2) + ((fCarga - 1) * 2))));

    const decision = construirDecisionEscenarioDecision_(oportunidades, carga, riesgo);
    rowsEsc.push([
      name,
      objetivo,
      Math.round(prob * 100) + '%',
      oportunidades,
      carga,
      riesgo,
      decision.decision,
      decision.paso
    ]);
  }

  if (rowsEsc.length) {
    sheet.getRange(14, 1, rowsEsc.length, 8).setValues(rowsEsc);
    sheet.getRange(14, 1, rowsEsc.length, 8)
      .setBackground('#FFF7ED')
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, true, true, '#FDBA74', SpreadsheetApp.BorderStyle.SOLID);

    for (let i = 0; i < rowsEsc.length; i++) {
      const rN = 14 + i;
      const decisionTxt = sanitizeValue_(rowsEsc[i][6]).toUpperCase();
      let color = '#FED7AA';
      if (decisionTxt.indexOf('ACELERAR') !== -1) color = '#BBF7D0';
      if (decisionTxt.indexOf('EQUILIBRAR') !== -1) color = '#FDE68A';
      if (decisionTxt.indexOf('PAUSAR') !== -1) color = '#FECACA';
      sheet.getRange(rN, 7, 1, 2).setBackground(color).setFontWeight('bold');
    }
  }

  sheet.autoResizeColumns(1, 8);
  sheet.setFrozenRows(2);
}

function construirDecisionEscenarioDecision_(oportunidades, carga, riesgo) {
  if (riesgo >= 8) {
    return {
      decision: 'PAUSAR Y ORDENAR DATOS',
      paso: 'Revisar fechas, contactos y links antes de ampliar pipeline.'
    };
  }

  if (oportunidades >= 5 && carga <= 7) {
    return {
      decision: 'ACELERAR CAPTACION',
      paso: 'Priorizar convocatorias ABIERTA con cierre mas cercano.'
    };
  }

  if (oportunidades <= 2) {
    return {
      decision: 'EXPANDIR RADAR',
      paso: 'Subir nuevas fuentes para aumentar oportunidades reales.'
    };
  }

  return {
    decision: 'EQUILIBRAR EJECUCION',
    paso: 'Mantener ritmo y reforzar solo oportunidades mejor perfiladas.'
  };
}

function asegurarBloqueosCalculados_(sheet, forceReset) {
  if (!sheet) return;

  const protectedRanges = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  let already = false;
  for (let i = 0; i < protectedRanges.length; i++) {
    const desc = sanitizeValue_(protectedRanges[i].getDescription());
    if (desc.indexOf(CRM_LOCK_PREFIX) === 0) {
      already = true;
      if (forceReset) {
        protectedRanges[i].remove();
      }
    }
  }

  if (already && !forceReset) return;

  const totalRows = Math.max(2, sheet.getMaxRows());
  for (let i = 0; i < CRM_DECISION.LOCKED_COLUMNS.length; i++) {
    const col = CRM_DECISION.LOCKED_COLUMNS[i];
    const rg = sheet.getRange(2, col, totalRows - 1, 1);
    const protection = rg.protect().setDescription(CRM_LOCK_PREFIX + '_COL_' + col);
    fijarPermisosProteccionDecision_(protection);
  }
}

function quitarBloqueosCalculados_(sheet) {
  if (!sheet) return;
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (let i = 0; i < protections.length; i++) {
    const desc = sanitizeValue_(protections[i].getDescription());
    if (desc.indexOf(CRM_LOCK_PREFIX) === 0) {
      protections[i].remove();
    }
  }
}

function bloquearPanelDecision_(sheet) {
  if (!sheet) return;
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  let found = null;
  for (let i = 0; i < protections.length; i++) {
    const desc = sanitizeValue_(protections[i].getDescription());
    if (desc === CRM_LOCK_PREFIX + '_PANEL') {
      found = protections[i];
      break;
    }
  }

  if (!found) {
    found = sheet.protect().setDescription(CRM_LOCK_PREFIX + '_PANEL');
  }
  fijarPermisosProteccionDecision_(found);
}

function quitarBloqueoPanelDecision_(sheet) {
  if (!sheet) return;
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (let i = 0; i < protections.length; i++) {
    const desc = sanitizeValue_(protections[i].getDescription());
    if (desc === CRM_LOCK_PREFIX + '_PANEL') {
      protections[i].remove();
    }
  }
}

function fijarPermisosProteccionDecision_(protection) {
  if (!protection) return;
  try {
    const editors = protection.getEditors();
    if (editors && editors.length) {
      protection.removeEditors(editors);
    }
    const me = Session.getEffectiveUser();
    const mail = me ? sanitizeValue_(me.getEmail()) : '';
    if (mail) {
      protection.addEditor(mail);
    }
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
    protection.setWarningOnly(false);
  } catch (err) {
    // Si no hay permisos para tocar editores, se mantiene la proteccion creada.
  }
}
