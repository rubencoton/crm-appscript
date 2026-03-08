// =============================================================================
// CRM FESTIVALES - HOMOGENEIDAD DE COLUMNAS + DISENO VISUAL
// Listo para Google Apps Script
// =============================================================================

const FEST_HOMO = {
  HEADER: [
    'NOMBRE FESTIVAL',
    'GENERO',
    'AFORO',
    'UBICACION',
    'PROVINCIA',
    'CCAA',
    'EMAIL',
    'TELEFONO',
    'NOMBRE CONTACTO',
    'OBSERVACIONES'
  ],
  COLORS: {
    HEADER_BG: '#8B0000',
    HEADER_FG: '#FFFFFF',
    BODY_BG: '#FFFFFF',
    BAND_BG: '#FAF3F3',
    WARN_BG: '#FFF4CC',
    ERROR_BG: '#FDE2E2',
    GRID: '#D9D9D9',
    LINK: '#1A73E8'
  }
};

/**
 * Si ya tienes un onOpen() en tu proyecto, llama a esta funcion desde ahi:
 * onOpenFestivalesHomogeneidad_();
 */
function onOpenFestivalesHomogeneidad_() {
  SpreadsheetApp.getUi()
    .createMenu('CRM Festivales')
    .addItem('Homogeneizar columnas + diseno', 'homogeneizarCRMFestivales')
    .addItem('Solo aplicar diseno', 'aplicarDisenoCRMFestivales')
    .addToUi();
}

/**
 * 1) Unifica TODAS las pestanas de festivales al mismo orden de columnas.
 * 2) Aplica diseno visual consistente para una lectura comoda.
 */
function homogeneizarCRMFestivales() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    SpreadsheetApp.getUi().alert('Hay otro proceso en curso. Intenta de nuevo en unos segundos.');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = getFestivalSheets_(ss);
    if (!sheets.length) {
      SpreadsheetApp.getUi().alert('No encontre pestanas de festivales para procesar.');
      return;
    }

    let totalRows = 0;
    let totalSheets = 0;

    sheets.forEach((sheet) => {
      const data = sheet.getDataRange().getValues();
      const normalizedRows = normalizeSheetRows_(data);
      rewriteSheet_(sheet, normalizedRows);
      applyVisualDesignToSheet_(sheet);
      totalRows += normalizedRows.length;
      totalSheets += 1;
    });

    SpreadsheetApp.getUi().alert(
      'Homogeneidad aplicada.\n' +
      'Pestanas procesadas: ' + totalSheets + '\n' +
      'Filas normalizadas: ' + totalRows + '\n\n' +
      'Columnas fijadas en A:J con el mismo orden para todas.'
    );
  } finally {
    lock.releaseLock();
  }
}

/**
 * Reaplica solo el formato visual a las pestanas detectadas.
 */
function aplicarDisenoCRMFestivales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getFestivalSheets_(ss);
  if (!sheets.length) {
    SpreadsheetApp.getUi().alert('No encontre pestanas de festivales para formatear.');
    return;
  }

  sheets.forEach((sheet) => applyVisualDesignToSheet_(sheet));
  SpreadsheetApp.getUi().alert('Diseno visual aplicado a ' + sheets.length + ' pestanas.');
}

function getFestivalSheets_(ss) {
  const valid = [];
  const reMain = /^(URBAN|POP|INDIE|ROCK|ELECTR|JAZZ|FLAM|RUMBA|MR|MC)_(S|L|XL)$/i;
  const rePending = /^PTE[_\-]/i;

  ss.getSheets().forEach((sheet) => {
    const name = (sheet.getName() || '').trim();
    if (reMain.test(name) || rePending.test(name)) {
      valid.push(sheet);
    }
  });

  return valid;
}

function normalizeSheetRows_(data) {
  if (!data || data.length === 0) return [];

  const header = data[0];
  const map = buildHeaderMap_(header);
  const out = [];

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const obj = {
      nombre: cleanText_(valueAt_(row, map.nombre)),
      genero: cleanText_(valueAt_(row, map.genero)),
      aforo: normalizeAforoForDisplay_(valueAt_(row, map.aforo)),
      ubicacion: cleanText_(valueAt_(row, map.ubicacion)),
      provincia: cleanText_(valueAt_(row, map.provincia)),
      ccaa: cleanText_(valueAt_(row, map.ccaa)),
      email: normalizeEmailCell_(valueAt_(row, map.email)),
      telefono: normalizePhoneCell_(valueAt_(row, map.telefono)),
      contacto: normalizeContactName_(valueAt_(row, map.contacto)),
      observaciones: cleanText_(valueAt_(row, map.observaciones))
    };

    // Descarta filas completamente vacias.
    if (!hasAnyValue_(obj)) continue;
    out.push([
      obj.nombre,
      obj.genero,
      obj.aforo,
      obj.ubicacion,
      obj.provincia,
      obj.ccaa,
      obj.email,
      obj.telefono,
      obj.contacto,
      obj.observaciones
    ]);
  }

  return out;
}

function rewriteSheet_(sheet, rows) {
  sheet.clearContents();

  sheet.getRange(1, 1, 1, FEST_HOMO.HEADER.length).setValues([FEST_HOMO.HEADER]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, FEST_HOMO.HEADER.length).setValues(rows);
  }
}

function applyVisualDesignToSheet_(sheet) {
  const lastRow = Math.max(2, sheet.getLastRow());
  const lastCol = FEST_HOMO.HEADER.length;

  // Limpia reglas anteriores para evitar acumulacion visual.
  sheet.setConditionalFormatRules([]);
  sheet.getBandings().forEach((b) => b.remove());

  // Cabecera corporativa.
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange
    .setBackground(FEST_HOMO.COLORS.HEADER_BG)
    .setFontColor(FEST_HOMO.COLORS.HEADER_FG)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, FEST_HOMO.COLORS.GRID, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Cuerpo.
  const bodyRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  bodyRange
    .setBackground(FEST_HOMO.COLORS.BODY_BG)
    .setFontColor('#222222')
    .setFontFamily('Roboto')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, FEST_HOMO.COLORS.GRID, SpreadsheetApp.BorderStyle.SOLID);

  // Zebra visual (bandas alternas) para mejor lectura.
  if (lastRow > 2) {
    const altRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const banding = altRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    banding.setFirstRowColor(FEST_HOMO.COLORS.BODY_BG);
    banding.setSecondRowColor(FEST_HOMO.COLORS.BAND_BG);
  }

  // Congela cabecera + filtro.
  sheet.setFrozenRows(1);
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  sheet.getRange(1, 1, lastRow, lastCol).createFilter();

  // Anchos de columna para una lectura comoda.
  sheet.setColumnWidth(1, 270); // Nombre
  sheet.setColumnWidth(2, 110); // Genero
  sheet.setColumnWidth(3, 90);  // Aforo
  sheet.setColumnWidth(4, 170); // Ubicacion
  sheet.setColumnWidth(5, 150); // Provincia
  sheet.setColumnWidth(6, 150); // CCAA
  sheet.setColumnWidth(7, 260); // Email
  sheet.setColumnWidth(8, 140); // Telefono
  sheet.setColumnWidth(9, 190); // Contacto
  sheet.setColumnWidth(10, 260); // Observaciones

  // Alineaciones por tipo de dato.
  if (lastRow > 1) {
    sheet.getRange(2, 3, lastRow - 1, 1).setHorizontalAlignment('center'); // Aforo
    sheet.getRange(2, 2, lastRow - 1, 1).setHorizontalAlignment('center'); // Genero
    sheet.getRange(2, 7, lastRow - 1, 2).setHorizontalAlignment('left');    // Email/Telefono
  }

  // Destaca filas sin datos de contacto (muy visual para accion rapida).
  const warnRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"";OR($G2="";$H2=""))')
    .setBackground(FEST_HOMO.COLORS.WARN_BG)
    .setRanges([sheet.getRange(2, 1, Math.max(1, lastRow - 1), lastCol)])
    .build();

  const errorRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"";$G2="";$H2="";$I2="")')
    .setBackground(FEST_HOMO.COLORS.ERROR_BG)
    .setRanges([sheet.getRange(2, 1, Math.max(1, lastRow - 1), lastCol)])
    .build();

  sheet.setConditionalFormatRules([warnRule, errorRule]);
}

function buildHeaderMap_(headerRow) {
  const map = {};
  for (let c = 0; c < headerRow.length; c++) {
    const h = normalizeHeader_(headerRow[c]);
    if (!h) continue;

    if (h.indexOf('NOMBRE FESTIVAL') > -1 && map.nombre === undefined) map.nombre = c;
    if (h === 'GENERO' && map.genero === undefined) map.genero = c;
    if (h === 'AFORO' && map.aforo === undefined) map.aforo = c;
    if ((h === 'UBICACION' || h === 'MUNICIPIO') && map.ubicacion === undefined) map.ubicacion = c;
    if (h === 'PROVINCIA' && map.provincia === undefined) map.provincia = c;
    if ((h === 'CCAA' || h.indexOf('COMUNIDAD') > -1) && map.ccaa === undefined) map.ccaa = c;
    if ((h === 'EMAIL' || h === 'E MAIL') && map.email === undefined) map.email = c;
    if (h === 'TELEFONO' && map.telefono === undefined) map.telefono = c;
    if ((h === 'NOMBRE CONTACTO' || h === 'CONTACTO') && map.contacto === undefined) map.contacto = c;
    if ((h === 'OBSERVACIONES' || h === 'NOTAS') && map.observaciones === undefined) map.observaciones = c;
  }

  // Fallback por posicion si alguna cabecera viene rota.
  if (map.nombre === undefined) map.nombre = 0;
  if (map.genero === undefined) map.genero = 1;
  if (map.aforo === undefined) map.aforo = 2;
  if (map.ubicacion === undefined) map.ubicacion = 3;
  if (map.provincia === undefined) map.provincia = 4;
  if (map.ccaa === undefined) map.ccaa = 5;
  if (map.email === undefined) map.email = 6;
  if (map.telefono === undefined) map.telefono = 7;
  if (map.contacto === undefined) map.contacto = 8;
  if (map.observaciones === undefined) map.observaciones = 9;

  return map;
}

function normalizeHeader_(v) {
  return cleanText_(v)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function valueAt_(row, idx) {
  if (idx === undefined || idx < 0 || idx >= row.length) return '';
  return row[idx];
}

function cleanText_(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/\s+/g, ' ').trim();
}

function hasAnyValue_(obj) {
  return !!(
    obj.nombre ||
    obj.genero ||
    obj.aforo ||
    obj.ubicacion ||
    obj.provincia ||
    obj.ccaa ||
    obj.email ||
    obj.telefono ||
    obj.contacto ||
    obj.observaciones
  );
}

function normalizeAforoForDisplay_(v) {
  if (v === null || v === undefined || v === '') return '';
  const n = parseAforo_(v);
  return n === '' ? cleanText_(v) : n;
}

function parseAforo_(v) {
  if (typeof v === 'number' && !isNaN(v)) return Math.round(v);
  const txt = cleanText_(v);
  if (!txt) return '';
  const digits = txt.replace(/[^\d]/g, '');
  if (!digits) return '';
  return parseInt(digits, 10);
}

function normalizeEmailCell_(v) {
  const txt = cleanText_(v).toLowerCase();
  if (!txt) return '';

  const tokens = txt
    .split(/[;,|\/\s]+/)
    .map((x) => x.trim())
    .filter((x) => x);

  const valid = [];
  const seen = {};
  for (let i = 0; i < tokens.length; i++) {
    if (isValidEmail_(tokens[i]) && !seen[tokens[i]]) {
      seen[tokens[i]] = true;
      valid.push(tokens[i]);
    }
  }

  if (valid.length) return valid.join('; ');
  return txt;
}

function isValidEmail_(email) {
  return /^[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}$/i.test(email || '');
}

function normalizePhoneCell_(v) {
  const txt = cleanText_(v);
  if (!txt) return '';

  const candidates = txt
    .split(/[;,|\/]+/)
    .map((x) => x.trim())
    .filter((x) => x);

  for (let i = 0; i < candidates.length; i++) {
    const formatted = formatSpanishPhone_(candidates[i]);
    if (formatted) return formatted;
  }

  return txt;
}

function formatSpanishPhone_(raw) {
  const digits = cleanText_(raw).replace(/[^\d]/g, '');
  if (!digits) return '';

  let base = digits;
  if (base.indexOf('0034') === 0) base = base.substring(4);
  if (base.indexOf('34') === 0 && base.length === 11) base = base.substring(2);

  if (base.length !== 9) return '';
  if (!/^[6789]/.test(base)) return '';

  return '+34 ' + base.substring(0, 3) + ' ' + base.substring(3, 6) + ' ' + base.substring(6);
}

function normalizeContactName_(v) {
  const txt = cleanText_(v);
  if (!txt) return '';

  const lower = txt.toLowerCase();
  const keepLower = { de: true, del: true, la: true, las: true, los: true, y: true, e: true };
  const parts = lower.split(' ');
  const out = [];

  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    if (!p) continue;
    if (i > 0 && keepLower[p]) {
      out.push(p);
      continue;
    }
    out.push(p.charAt(0).toUpperCase() + p.slice(1));
  }

  return out.join(' ');
}

