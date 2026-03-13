function codexAuditarHojaExtrema(spreadsheetId) {
  var sid = spreadsheetId || '1LgZG2ObSCJzEQvrysu1NFFEvYlupLXVByDnIMCr-wYA';
  var ss = SpreadsheetApp.openById(sid);
  var sheets = ss.getSheets();
  var findings = [];
  var sheetSummaries = [];
  var totalProtected = 0;

  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var dr = sh.getDataRange();
    var rows = dr.getNumRows();
    var cols = dr.getNumColumns();
    var values = dr.getDisplayValues();
    var formulas = dr.getFormulas();
    var numberFormats = dr.getNumberFormats();
    var fontFamilies = dr.getFontFamilies();
    var fontStyles = dr.getFontStyles();
    var fontWeights = dr.getFontWeights();
    var hAlign = dr.getHorizontalAlignments();
    var vAlign = dr.getVerticalAlignments();
    var backgrounds = dr.getBackgrounds();
    var validations = dr.getDataValidations();
    var merges = dr.getMergedRanges();

    var numFmtCount = {};
    var fontCount = {};
    var bgCount = {};
    var hAlignCount = {};
    var vAlignCount = {};
    var valCount = 0;
    var valByCol = {};
    var formulaCount = 0;
    var errorCells = 0;
    var boldCount = 0;
    var italicCount = 0;
    var formulaSamples = [];

    for (var r = 0; r < rows; r++) {
      for (var c = 0; c < cols; c++) {
        var nf = numberFormats[r][c] || '';
        var ff = fontFamilies[r][c] || '';
        var bg = backgrounds[r][c] || '';
        var ha = hAlign[r][c] || 'UNSPECIFIED';
        var va = vAlign[r][c] || 'UNSPECIFIED';
        var f = formulas[r][c] || '';
        var dv = validations[r][c];
        var displayed = values[r][c] || '';

        numFmtCount[nf] = (numFmtCount[nf] || 0) + 1;
        fontCount[ff] = (fontCount[ff] || 0) + 1;
        bgCount[bg] = (bgCount[bg] || 0) + 1;
        hAlignCount[ha] = (hAlignCount[ha] || 0) + 1;
        vAlignCount[va] = (vAlignCount[va] || 0) + 1;

        if (fontWeights[r][c] === 'bold') boldCount++;
        if (fontStyles[r][c] === 'italic') italicCount++;

        if (dv) {
          valCount++;
          var colKey = String(c + 1);
          valByCol[colKey] = (valByCol[colKey] || 0) + 1;
        }

        if (f) {
          formulaCount++;
          if (formulaSamples.length < 20) {
            formulaSamples.push({
              cell: _auditCol_(c + 1) + String(r + 1),
              formula: f
            });
          }
        }
        if (String(displayed).indexOf('#') === 0) errorCells++;
      }
    }

    var condRules = sh.getConditionalFormatRules();
    var filterObj = sh.getFilter();
    var basicFilterA1 = filterObj ? filterObj.getRange().getA1Notation() : '';

    var protRange = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
    var protSheet = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET) || [];
    totalProtected += protRange.length + protSheet.length;

    var protPreview = [];
    for (var p1 = 0; p1 < protRange.length && protPreview.length < 10; p1++) {
      var pr = protRange[p1];
      protPreview.push({
        type: 'RANGE',
        description: pr.getDescription() || '',
        warningOnly: pr.isWarningOnly(),
        rangeA1: pr.getRange() ? pr.getRange().getA1Notation() : '',
        editorsCount: _safeEditorsCount_(pr)
      });
    }
    for (var p2 = 0; p2 < protSheet.length && protPreview.length < 10; p2++) {
      var ps = protSheet[p2];
      protPreview.push({
        type: 'SHEET',
        description: ps.getDescription() || '',
        warningOnly: ps.isWarningOnly(),
        rangeA1: sh.getName(),
        editorsCount: _safeEditorsCount_(ps)
      });
    }

    var previewHeader = values.length ? values[0].slice(0, Math.min(cols, 25)) : [];
    var sampleRows = [];
    for (var sr = 1; sr < Math.min(rows, 6); sr++) {
      sampleRows.push(values[sr].slice(0, Math.min(cols, 25)));
    }

    sheetSummaries.push({
      title: sh.getName(),
      index: sh.getIndex(),
      hidden: sh.isSheetHidden(),
      usedRange: dr.getA1Notation(),
      usedRows: rows,
      usedCols: cols,
      maxRows: sh.getMaxRows(),
      maxCols: sh.getMaxColumns(),
      frozenRows: sh.getFrozenRows(),
      frozenCols: sh.getFrozenColumns(),
      dataPreview: {
        header: previewHeader,
        sampleRows: sampleRows
      },
      formulas: {
        count: formulaCount,
        samples: formulaSamples
      },
      formats: {
        numberFormats: _topMap_(numFmtCount, 15),
        fontFamilies: _topMap_(fontCount, 10),
        backgrounds: _topMap_(bgCount, 10),
        hAlign: _topMap_(hAlignCount, 10),
        vAlign: _topMap_(vAlignCount, 10),
        boldCount: boldCount,
        italicCount: italicCount
      },
      merges: {
        count: merges.length,
        samples: _mergeSamples_(merges, 20)
      },
      validations: {
        count: valCount,
        byCol: valByCol
      },
      filters: {
        basicFilterRange: basicFilterA1,
        filterViews: 'NOT_AVAILABLE_IN_SPREADSHEETAPP'
      },
      conditionalFormatting: {
        rulesCount: condRules.length
      },
      protections: {
        rangeCount: protRange.length,
        sheetCount: protSheet.length,
        samples: protPreview
      },
      cellErrorsCount: errorCells
    });
  }

  var concursos = _findSheetSummary_(sheetSummaries, 'CONCURSOS');
  if (!concursos) {
    _pushFinding_(findings, 'CRITICAL', 'SHEET_STRUCTURE', 'No existe pestaña CONCURSOS',
      'No se puede validar la matriz principal.', {}, 'Crear o renombrar la pestaña principal a CONCURSOS.');
  } else {
    _auditConcursos_(concursos, findings);
  }

  for (var i = 0; i < sheetSummaries.length; i++) {
    if (sheetSummaries[i].cellErrorsCount > 0) {
      _pushFinding_(findings, 'HIGH', 'SHEET_FORMULAS', 'Celdas con error visible',
        'Se detectaron celdas con #ERROR/#REF/etc.', {
          sheet: sheetSummaries[i].title,
          errorCells: sheetSummaries[i].cellErrorsCount
        }, 'Corregir formulas antes de automatizar procesos.');
    }
  }

  if (totalProtected === 0) {
    _pushFinding_(findings, 'LOW', 'SHEET_PROTECTION', 'Sin protecciones',
      'No hay protecciones de rango ni hoja.', {}, 'Proteger columnas calculadas y cabeceras.');
  }

  var report = {
    meta: {
      timestamp: new Date().toISOString(),
      spreadsheetId: sid,
      spreadsheetTitle: ss.getName(),
      timeZone: ss.getSpreadsheetTimeZone(),
      locale: ss.getSpreadsheetLocale()
    },
    coverage: {
      structure: true,
      visibleDataAndFormulas: true,
      formats: true,
      mergedCells: true,
      validations: true,
      filtersAndViews: true,
      conditionalFormatting: true,
      protections: true,
      filterViewsNote: 'Filter Views detalladas no expuestas por SpreadsheetApp; solo filtro basico.'
    },
    sheetCount: sheetSummaries.length,
    sheets: sheetSummaries,
    findings: findings
  };

  return JSON.stringify(report);
}

function _auditConcursos_(sheetSummary, findings) {
  var header = sheetSummary.dataPreview.header || [];
  var rows = sheetSummary.dataPreview.sampleRows || [];
  if (!header.length) {
    _pushFinding_(findings, 'CRITICAL', 'SHEET_DATA', 'CONCURSOS sin cabecera',
      'No hay fila de cabecera visible.', {}, 'Revisar rango usado y cabeceras.');
    return;
  }

  var hnorm = [];
  for (var i = 0; i < header.length; i++) hnorm.push(_auditNorm_(header[i]));
  var hm = {};
  for (var j = 0; j < hnorm.length; j++) hm[hnorm[j]] = j;

  var expected = ['NOMBRE CONCURSO', 'ESTADO', 'INSCRIPCION', 'FECHA LIMITE INSCRIPCION', 'EMAIL'];
  var missing = [];
  for (var e = 0; e < expected.length; e++) {
    if (hm[expected[e]] === undefined) missing.push(expected[e]);
  }
  if (missing.length) {
    _pushFinding_(findings, 'HIGH', 'SHEET_STRUCTURE', 'Cabeceras ausentes en CONCURSOS',
      'Faltan columnas clave para el motor.', { missing: missing }, 'Alinear cabeceras exactas.');
  }

  var idxEstado = hm['ESTADO'] !== undefined ? hm['ESTADO'] + 1 : 2;
  var idxIns = hm['INSCRIPCION'] !== undefined ? hm['INSCRIPCION'] + 1 : 3;
  var validB = Number(sheetSummary.validations.byCol[String(idxEstado)] || 0);
  var validC = Number(sheetSummary.validations.byCol[String(idxIns)] || 0);
  var dataRows = Math.max(0, sheetSummary.usedRows - 1);
  if (dataRows > 0) {
    var threshold = Math.floor(dataRows * 0.8);
    var weak = [];
    if (validB < threshold) weak.push({ column: 'ESTADO', colIndex: idxEstado, validatedCells: validB, dataRows: dataRows });
    if (validC < threshold) weak.push({ column: 'INSCRIPCION', colIndex: idxIns, validatedCells: validC, dataRows: dataRows });
    if (weak.length) {
      _pushFinding_(findings, 'HIGH', 'SHEET_VALIDATIONS', 'Validaciones incompletas en CONCURSOS',
        'Las columnas principales no tienen validacion consistente.', { columns: weak }, 'Aplicar dropdown a todas las filas activas.');
    }
  }

  var mojibake = 0;
  var allText = JSON.stringify(header) + JSON.stringify(rows);
  if (allText.indexOf('ðŸ') >= 0 || allText.indexOf('Ã') >= 0 || allText.indexOf('Â') >= 0) {
    mojibake = 1;
  }
  if (mojibake) {
    _pushFinding_(findings, 'MEDIUM', 'SHEET_UX', 'Texto con codificacion rota detectado',
      'Se detectaron tokens corruptos en datos visibles.', {}, 'Normalizar codificacion y limpiar literales corruptos.');
  }
}

function _findSheetSummary_(arr, name) {
  var needle = _auditNorm_(name);
  for (var i = 0; i < arr.length; i++) {
    if (_auditNorm_(arr[i].title) === needle) return arr[i];
  }
  return null;
}

function _mergeSamples_(ranges, limit) {
  var out = [];
  for (var i = 0; i < ranges.length && out.length < limit; i++) {
    out.push(ranges[i].getA1Notation());
  }
  return out;
}

function _topMap_(m, limit) {
  var arr = [];
  for (var k in m) {
    if (!m.hasOwnProperty(k)) continue;
    arr.push({ key: k, count: m[k] });
  }
  arr.sort(function(a, b) { return b.count - a.count; });
  var out = {};
  for (var i = 0; i < arr.length && i < limit; i++) out[arr[i].key] = arr[i].count;
  return out;
}

function _safeEditorsCount_(protection) {
  try {
    return protection.getEditors().length;
  } catch (e) {
    return -1;
  }
}

function _auditCol_(n) {
  var s = '';
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s || 'A';
}

function _auditNorm_(x) {
  var s = String(x || '').trim().toUpperCase();
  s = s.replace(/Á/g, 'A').replace(/É/g, 'E').replace(/Í/g, 'I').replace(/Ó/g, 'O').replace(/Ú/g, 'U').replace(/Ü/g, 'U').replace(/Ñ/g, 'N');
  s = s.replace(/\s+/g, ' ');
  return s;
}

function _pushFinding_(arr, severity, area, title, detail, evidence, recommendation) {
  arr.push({
    severity: severity,
    area: area,
    title: title,
    detail: detail,
    evidence: evidence || {},
    recommendation: recommendation || ''
  });
}
