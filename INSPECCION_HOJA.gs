function codexInspeccionarHojaVinculada() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No hay hoja activa vinculada al script.');
  }

  const report = {
    generatedAt: new Date().toISOString(),
    workbook: {
      id: ss.getId(),
      name: ss.getName(),
      url: ss.getUrl(),
      timeZone: ss.getSpreadsheetTimeZone(),
      locale: ss.getSpreadsheetLocale(),
      sheetsCount: ss.getSheets().length
    },
    permissions: getWorkbookPermissions_(ss),
    sheets: ss.getSheets().map((sheet) => inspectSheet_(sheet))
  };

  const file = DriveApp.getFolderById(DriveApp.getRootFolder().getId())
    .createFile('codex_inspeccion_' + ss.getId() + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.json', JSON.stringify(report, null, 2), MimeType.PLAIN_TEXT);

  return {
    ok: true,
    reportSummary: {
      workbook: report.workbook,
      permissions: report.permissions,
      sheetNames: report.sheets.map((s) => s.name),
      outputFileId: file.getId(),
      outputFileUrl: file.getUrl()
    }
  };
}

function inspectSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();
  const usedRows = Math.max(lastRow, 1);
  const usedCols = Math.max(lastColumn, 1);
  const usedRange = sheet.getRange(1, 1, usedRows, usedCols);

  const values = usedRange.getDisplayValues();
  const formulas = usedRange.getFormulas();
  const numberFormats = usedRange.getNumberFormats();
  const backgrounds = usedRange.getBackgrounds();
  const fontFamilies = usedRange.getFontFamilies();
  const fontSizes = usedRange.getFontSizes();
  const fontWeights = usedRange.getFontWeights();
  const fontStyles = usedRange.getFontStyles();
  const horizontalAlignments = usedRange.getHorizontalAlignments();
  const verticalAlignments = usedRange.getVerticalAlignments();

  const formulaCells = [];
  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < formulas[r].length; c++) {
      const formula = formulas[r][c];
      if (formula) {
        formulaCells.push({
          a1: sheet.getRange(r + 1, c + 1).getA1Notation(),
          formula
        });
      }
    }
  }

  const validations = [];
  const dv = usedRange.getDataValidations();
  for (let r = 0; r < dv.length; r++) {
    for (let c = 0; c < dv[r].length; c++) {
      const rule = dv[r][c];
      if (!rule) continue;
      validations.push({
        a1: sheet.getRange(r + 1, c + 1).getA1Notation(),
        criteriaType: String(rule.getCriteriaType()),
        allowInvalid: rule.getAllowInvalid(),
        helpText: rule.getHelpText(),
        criteriaValues: (rule.getCriteriaValues() || []).map((v) => {
          if (v && typeof v.getA1Notation === 'function') return v.getA1Notation();
          if (v instanceof Date) return v.toISOString();
          return String(v);
        })
      });
    }
  }

  const conditionalRules = sheet.getConditionalFormatRules().map((rule, idx) => ({
    index: idx + 1,
    ranges: rule.getRanges().map((rg) => rg.getA1Notation()),
    booleanCondition: normalizeBooleanCondition_(rule.getBooleanCondition()),
    gradientCondition: normalizeGradientCondition_(rule.getGradientCondition())
  }));

  const protectionsRange = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).map((p) => normalizeProtection_(p));
  const protectionsSheet = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).map((p) => normalizeProtection_(p));

  const mergedRanges = usedRange.getMergedRanges().map((rg) => rg.getA1Notation());

  const basicFilter = sheet.getFilter();
  const filter = basicFilter ? {
    range: basicFilter.getRange().getA1Notation(),
    criteriaColumns: collectFilterCriteria_(basicFilter, lastColumn)
  } : null;

  return {
    name: sheet.getName(),
    sheetId: sheet.getSheetId(),
    hidden: sheet.isSheetHidden(),
    tabColor: sheet.getTabColor(),
    frozenRows: sheet.getFrozenRows(),
    frozenColumns: sheet.getFrozenColumns(),
    dimensions: {
      maxRows,
      maxColumns,
      usedRows: lastRow,
      usedColumns: lastColumn,
      usedRangeA1: toA1Range_(usedRows, usedCols)
    },
    visibleDataSample: {
      header: values[0] || [],
      rows: values.slice(1, 51),
      rowCountIncluded: Math.max(values.length - 1, 0)
    },
    formulas: {
      total: formulaCells.length,
      sample: formulaCells.slice(0, 400)
    },
    formatting: {
      backgroundsUnique: collectUniques_(backgrounds),
      fontFamiliesUnique: collectUniques_(fontFamilies),
      fontSizesUnique: collectUniques_(fontSizes),
      fontWeightsUnique: collectUniques_(fontWeights),
      fontStylesUnique: collectUniques_(fontStyles),
      horizontalAlignmentsUnique: collectUniques_(horizontalAlignments),
      verticalAlignmentsUnique: collectUniques_(verticalAlignments),
      numberFormatsUnique: collectUniques_(numberFormats)
    },
    merges: mergedRanges,
    dataValidations: {
      total: validations.length,
      sample: validations.slice(0, 500)
    },
    filter,
    conditionalFormatting: {
      total: conditionalRules.length,
      rules: conditionalRules
    },
    protections: {
      sheet: protectionsSheet,
      ranges: protectionsRange
    },
    notes: [
      'Bordes y vistas de filtro no tienen getter completo en SpreadsheetApp. Para detalle total hay que usar Sheets API avanzada.'
    ]
  };
}

function normalizeBooleanCondition_(condition) {
  if (!condition) return null;
  return {
    type: String(condition.getCriteriaType()),
    values: (condition.getCriteriaValues() || []).map((v) => {
      if (v instanceof Date) return v.toISOString();
      return String(v);
    })
  };
}

function normalizeGradientCondition_(condition) {
  if (!condition) return null;
  const min = condition.getMinpoint();
  const mid = condition.getMidpoint();
  const max = condition.getMaxpoint();
  return {
    min: min ? {
      type: String(min.getType()),
      value: min.getValue(),
      color: min.getColor()
    } : null,
    mid: mid ? {
      type: String(mid.getType()),
      value: mid.getValue(),
      color: mid.getColor()
    } : null,
    max: max ? {
      type: String(max.getType()),
      value: max.getValue(),
      color: max.getColor()
    } : null
  };
}

function normalizeProtection_(protection) {
  return {
    description: protection.getDescription(),
    warningOnly: protection.isWarningOnly(),
    editors: protection.getEditors().map((u) => u.getEmail()),
    domainEdit: protection.canDomainEdit(),
    unprotectedRanges: typeof protection.getUnprotectedRanges === 'function'
      ? protection.getUnprotectedRanges().map((rg) => rg.getA1Notation())
      : []
  };
}

function getWorkbookPermissions_(ss) {
  const file = DriveApp.getFileById(ss.getId());
  return {
    owner: file.getOwner() ? file.getOwner().getEmail() : null,
    editors: file.getEditors().map((u) => u.getEmail()),
    viewers: file.getViewers().map((u) => u.getEmail()),
    sharingAccess: String(file.getSharingAccess()),
    sharingPermission: String(file.getSharingPermission())
  };
}

function collectUniques_(matrix) {
  const set = new Set();
  matrix.forEach((row) => row.forEach((v) => set.add(String(v))));
  return Array.from(set).slice(0, 200);
}

function toA1Range_(rows, cols) {
  return 'A1:' + columnToLetter_(cols) + rows;
}

function columnToLetter_(column) {
  let temp = '';
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function collectFilterCriteria_(filter, lastColumn) {
  const out = [];
  for (let c = 1; c <= lastColumn; c++) {
    const criteria = filter.getColumnFilterCriteria(c);
    if (!criteria) continue;
    out.push({
      column: c,
      hiddenValues: criteria.getHiddenValues(),
      visibleBackgroundColor: criteria.getVisibleBackgroundColor(),
      visibleForegroundColor: criteria.getVisibleForegroundColor(),
      criteriaType: criteria.getCriteriaType() ? String(criteria.getCriteriaType()) : null,
      criteriaValues: (criteria.getCriteriaValues() || []).map((v) => String(v))
    });
  }
  return out;
}