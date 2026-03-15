function codexEjecutarLimpiezaPendienteOnOpen_() {
  var targetSpreadsheetId = '1LgZG2ObSCJzEQvrysu1NFFEvYlupLXVByDnIMCr-wYA';
  var doneKey = 'CODEX_CLEANUP_ONOPEN_V2_DONE';
  var lastRunKey = 'CODEX_CLEANUP_ONOPEN_V2_LASTRUN_TS';
  var minIntervalMs = 6 * 60 * 60 * 1000;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  if (ss.getId() !== targetSpreadsheetId) return;

  var props = PropertiesService.getDocumentProperties();
  var done = props.getProperty(doneKey) === '1';
  var lastRunTs = Number(props.getProperty(lastRunKey) || 0);
  var nowTs = Date.now();
  if (done && lastRunTs > 0 && (nowTs - lastRunTs) < minIntervalMs) return;

  try {
    var reportJson = codexAplicarLimpiezaOperativaAyudas(ss.getId());
    props.setProperty('CODEX_CLEANUP_LAST_REPORT_JSON', String(reportJson || ''));
    props.setProperty(doneKey, '1');
    props.setProperty(lastRunKey, String(nowTs));
  } catch (err) {
    Logger.log('Error en auto-limpieza onOpen: ' + (err && err.message ? err.message : err));
  }
}

function codexForzarLimpiezaOperativaAyudas() {
  var reportJson = codexAplicarLimpiezaOperativaAyudas();
  var summary = 'Limpieza operativa ejecutada.';
  try {
    var report = JSON.parse(reportJson || '{}');
    var emails = report && report.steps && report.steps.emails ? report.steps.emails : {};
    var trimRows = report && report.steps && report.steps.trimRows ? report.steps.trimRows : {};
    var residual = report && report.steps && report.steps.residualTabs ? report.steps.residualTabs : {};
    summary =
      'Limpieza operativa completada\n' +
      '- Emails normalizados: ' + Number(emails.changedEmails || 0) + '\n' +
      '- Emails extra movidos a notas: ' + Number(emails.extraEmailsMovedToNotes || 0) + '\n' +
      '- Filas eliminadas: ' + Number(trimRows.deletedRows || 0) + '\n' +
      '- Pestanas residuales eliminadas: ' + Number(residual.deletedCount || 0);
  } catch (e) {
    summary = 'Limpieza operativa ejecutada, pero no se pudo resumir el reporte.';
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(summary, 'Codex Cleanup', 8);
  return reportJson;
}

function codexAplicarLimpiezaOperativaAyudas(spreadsheetId) {
  var defaultId = '1LgZG2ObSCJzEQvrysu1NFFEvYlupLXVByDnIMCr-wYA';
  var active = SpreadsheetApp.getActiveSpreadsheet();
  var ss = null;

  if (spreadsheetId) {
    if (active && active.getId() === spreadsheetId) {
      ss = active;
    } else {
      ss = SpreadsheetApp.openById(spreadsheetId);
    }
  } else if (active) {
    ss = active;
  } else {
    ss = SpreadsheetApp.openById(defaultId);
  }

  var sid = ss.getId();
  var concursos = ss.getSheetByName('CONCURSOS');
  if (!concursos) throw new Error('No existe la pestana CONCURSOS.');

  var report = {
    ok: true,
    spreadsheetId: sid,
    spreadsheetTitle: ss.getName(),
    timestamp: new Date().toISOString(),
    steps: {
      emails: {},
      trimRows: {},
      residualTabs: {}
    }
  };

  report.steps.emails = _codexLimpiarEmailsConcursos_(concursos);
  report.steps.trimRows = _codexRecortarFilasSobrantes_(concursos);
  report.steps.residualTabs = _codexLimpiarPestanasResiduales_(ss);

  return JSON.stringify(report);
}

function _codexLimpiarEmailsConcursos_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var idxEmail = _codexFindHeaderIndex_(headers, ['EMAIL', 'CORREO', 'E-MAIL', 'MAIL']);
  var idxNotas = _codexFindHeaderIndex_(headers, ['NOTAS', 'OBSERVACIONES', 'OBS']);

  if (idxEmail < 1) {
    throw new Error('No se encontro columna EMAIL/CORREO en CONCURSOS.');
  }

  var rows = Math.max(0, lastRow - 1);
  if (!rows) {
    return {
      sheet: sheet.getName(),
      columnEmail: idxEmail,
      columnNotas: idxNotas,
      scannedRows: 0,
      changedEmails: 0,
      clearedPlaceholders: 0,
      clearedInvalid: 0,
      multiEmailCells: 0,
      extraEmailsMovedToNotes: 0
    };
  }

  var emailRange = sheet.getRange(2, idxEmail, rows, 1);
  var emailValues = emailRange.getDisplayValues();
  var noteValues = null;
  if (idxNotas > 0) {
    noteValues = sheet.getRange(2, idxNotas, rows, 1).getDisplayValues();
  }

  var changedEmails = 0;
  var clearedPlaceholders = 0;
  var clearedInvalid = 0;
  var multiEmailCells = 0;
  var extraEmailsMovedToNotes = 0;
  var changedNotes = 0;

  for (var r = 0; r < rows; r++) {
    var original = emailValues[r][0];
    var normalized = _codexNormalizeEmailCell_(original);

    if (normalized.wasPlaceholder) clearedPlaceholders++;
    if (normalized.invalidInput) clearedInvalid++;
    if (normalized.hadMultiple) multiEmailCells++;

    if (normalized.changed) {
      emailValues[r][0] = normalized.primary;
      changedEmails++;
    } else {
      emailValues[r][0] = normalized.primary;
    }

    if (idxNotas > 0 && normalized.extras.length) {
      var currentNote = _codexSan_(noteValues[r][0]);
      var extraText = 'Emails extra detectados: ' + normalized.extras.join(', ');
      if (currentNote.indexOf(extraText) === -1) {
        noteValues[r][0] = currentNote ? (currentNote + ' | ' + extraText) : extraText;
        changedNotes++;
        extraEmailsMovedToNotes += normalized.extras.length;
      }
    }
  }

  emailRange.setValues(emailValues);
  if (idxNotas > 0 && changedNotes > 0) {
    sheet.getRange(2, idxNotas, rows, 1).setValues(noteValues);
  }

  return {
    sheet: sheet.getName(),
    columnEmail: idxEmail,
    columnNotas: idxNotas,
    scannedRows: rows,
    changedEmails: changedEmails,
    clearedPlaceholders: clearedPlaceholders,
    clearedInvalid: clearedInvalid,
    multiEmailCells: multiEmailCells,
    extraEmailsMovedToNotes: extraEmailsMovedToNotes
  };
}

function _codexNormalizeEmailCell_(raw) {
  var source = _codexSan_(raw);
  if (!source) {
    return {
      primary: '',
      extras: [],
      changed: false,
      wasPlaceholder: false,
      invalidInput: false,
      hadMultiple: false
    };
  }

  var lowered = source.toLowerCase();
  var placeholders = {
    'n/a': true,
    'na': true,
    'none': true,
    'null': true,
    's/d': true,
    's/n': true,
    'no aplica': true,
    'pendiente': true,
    'sin email': true,
    'sin correo': true,
    'no email': true,
    'no correo': true,
    'no hay': true,
    'no disponible': true,
    'no publicado': true,
    'no localizable': true,
    'desconocido': true,
    '-': true,
    '--': true
  };

  if (placeholders[lowered]) {
    return {
      primary: '',
      extras: [],
      changed: source !== '',
      wasPlaceholder: true,
      invalidInput: false,
      hadMultiple: false
    };
  }

  var compact = source
    .replace(/[\n\r\t]+/g, ' ')
    .replace(/(\(|\[)?\s*arroba\s*(\)|\])?/ig, '@')
    .replace(/\s*(\(at\)|\[at\])\s*/ig, '@')
    .replace(/\s+(dot|punto)\s+/ig, '.')
    .replace(/\s*@\s*/g, '@')
    .replace(/\s*\.\s*/g, '.')
    .replace(/[;|/]+/g, ',')
    .replace(/\s+y\s+/ig, ',')
    .replace(/\s+/g, ' ');
  var matches = compact.match(/[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}/ig) || [];
  var unique = [];
  var seen = {};
  for (var i = 0; i < matches.length; i++) {
    var m = String(matches[i] || '').toLowerCase();
    if (m && !seen[m]) {
      seen[m] = true;
      unique.push(m);
    }
  }

  if (!unique.length) {
    return {
      primary: '',
      extras: [],
      changed: source !== '',
      wasPlaceholder: false,
      invalidInput: true,
      hadMultiple: false
    };
  }

  var primary = unique[0];
  var extras = unique.slice(1);
  return {
    primary: primary,
    extras: extras,
    changed: primary !== source,
    wasPlaceholder: false,
    invalidInput: false,
    hadMultiple: unique.length > 1
  };
}

function _codexRecortarFilasSobrantes_(sheet) {
  var lastCol = Math.max(1, sheet.getLastColumn());
  var lastRow = Math.max(1, sheet.getLastRow());
  var values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var formulas = sheet.getRange(1, 1, lastRow, lastCol).getFormulas();
  var lastMeaningfulRow = 1;

  for (var r = lastRow - 1; r >= 0; r--) {
    var hasData = false;
    for (var c = 0; c < lastCol; c++) {
      if (_codexSan_(values[r][c]) || _codexSan_(formulas[r][c])) {
        hasData = true;
        break;
      }
    }
    if (hasData) {
      lastMeaningfulRow = r + 1;
      break;
    }
  }

  var bufferRows = 50;
  var minRows = 160;
  var targetRows = Math.max(minRows, lastMeaningfulRow + bufferRows);
  var beforeRows = sheet.getMaxRows();
  var deleted = 0;

  if (beforeRows > targetRows) {
    deleted = beforeRows - targetRows;
    sheet.deleteRows(targetRows + 1, deleted);
  }

  return {
    sheet: sheet.getName(),
    lastMeaningfulRow: lastMeaningfulRow,
    bufferRows: bufferRows,
    minRows: minRows,
    maxRowsBefore: beforeRows,
    maxRowsAfter: sheet.getMaxRows(),
    deletedRows: deleted
  };
}

function _codexLimpiarPestanasResiduales_(ss) {
  var keep = {
    CONCURSOS: true,
    'NUEVOS CONCURSOS': true,
    BANDAS: true,
    CORREO: true,
    ESCENARIOS: true,
    PANEL_DECISION: true
  };

  var sheets = ss.getSheets();
  var deleted = [];
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var name = sh.getName();
    var norm = _codexNorm_(name);
    if (keep[norm]) continue;
    if (!_codexPestanaResidualPorNombre_(name)) continue;
    if (_codexSheetHasRealData_(sh)) continue;
    if (ss.getSheets().length <= 1) break;
    ss.deleteSheet(sh);
    deleted.push(name);
  }

  return {
    deletedCount: deleted.length,
    deletedNames: deleted
  };
}

function _codexPestanaResidualPorNombre_(name) {
  var n = _codexSan_(name);
  if (!n) return false;
  if (/^Hoja\s*\d+$/i.test(n)) return true;
  if (/^Hoja$/i.test(n)) return true;
  if (/^Sheet\s*\d*$/i.test(n)) return true;
  return false;
}

function _codexSheetHasRealData_(sheet) {
  var lr = Math.max(1, sheet.getLastRow());
  var lc = Math.max(1, sheet.getLastColumn());
  var rows = Math.min(lr, 200);
  var cols = Math.min(lc, 20);
  var vals = sheet.getRange(1, 1, rows, cols).getDisplayValues();
  var forms = sheet.getRange(1, 1, rows, cols).getFormulas();
  var nonEmpty = 0;
  var nonEmptyRows = 0;
  var firstRowNonEmpty = 0;

  for (var r = 0; r < rows; r++) {
    var rowHasData = false;
    var rowNonEmpty = 0;
    for (var c = 0; c < cols; c++) {
      if (_codexSan_(vals[r][c]) || _codexSan_(forms[r][c])) {
        nonEmpty++;
        rowHasData = true;
        rowNonEmpty++;
      }
    }
    if (r === 0) firstRowNonEmpty = rowNonEmpty;
    if (rowHasData) nonEmptyRows++;
  }
  if (nonEmpty <= 3) return false;

  // Cabecera suelta duplicada (sin filas reales) se considera residual.
  if (nonEmptyRows === 1 && firstRowNonEmpty >= 6 && _codexLooksLikeConcursosHeader_(vals[0])) {
    return false;
  }

  return true;
}

function _codexLooksLikeConcursosHeader_(rowValues) {
  if (!rowValues || !rowValues.length) return false;
  var expected = {
    'NOMBRE CONCURSO': true,
    ESTADO: true,
    INSCRIPCION: true,
    'FECHA LIMITE INSCRIPCION': true,
    'LINK BASES/FORMULARIO': true,
    EMAIL: true
  };
  var matches = 0;
  for (var i = 0; i < rowValues.length; i++) {
    var key = _codexNorm_(rowValues[i]);
    if (expected[key]) matches++;
  }
  return matches >= 4;
}

function _codexFindHeaderIndex_(headers, candidates) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[_codexNorm_(headers[i])] = i + 1;
  }
  for (var j = 0; j < candidates.length; j++) {
    var key = _codexNorm_(candidates[j]);
    if (map[key]) return map[key];
  }
  return -1;
}

function _codexNorm_(value) {
  var s = _codexSan_(value).toUpperCase();
  if (s.normalize) {
    s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

function _codexSan_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

