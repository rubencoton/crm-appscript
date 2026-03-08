const fs = require('fs');
const path = 'C:/Users/elrub/Desktop/CARPETA CODEX/01_PROYECTOS/festivales-github/Código.js';
let t = fs.readFileSync(path, 'utf8');

function replaceOne(regex, replacement, label) {
  const next = t.replace(regex, replacement);
  if (next === t) throw new Error('No se pudo aplicar: ' + label);
  t = next;
}

replaceOne(
  /const FEST_MAX_RUNTIME_MS = 4\.7 \* 60 \* 1000;/,
  `const FEST_MAX_RUNTIME_MS = 4.7 * 60 * 1000;
const FEST_ARCHITECT = 'RUBEN COTON';
const FEST_GENRE_ORDER = ['URBAN', 'POP', 'INDIE', 'ROCK', 'ELECTR', 'JAZZ', 'FLAM', 'RUMBA', 'MEC', 'MFR'];
const FEST_SIZE_ORDER = ['S', 'L', 'XL'];
const FEST_GENRE_DROPDOWN = [
  '🧢 URBAN', '🎤 POP', '🎸 INDIE', '🤘 ROCK', '🎛️ ELECTR',
  '🎷 JAZZ', '💃 FLAM', '🪘 RUMBA', '🎼 MEC', '🌄 MFR'
];`,
  'constantes extra'
);

replaceOne(
  /function crearMenuCRMFestivales_\(\) \{[\s\S]*?\n\}/,
  `function crearMenuCRMFestivales_() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 CRM FESTIVALES | RUBEN COTON')
    .addItem('🚀 Escaner total + homogeneizar (seguro)', 'menuEscanerTotalCRMFestivales')
    .addItem('🎨 Solo armonizar diseno visual (seguro)', 'menuAplicarDisenoCRMFestivales')
    .addItem('📧 Depurar contactos local (seguro)', 'menuDepurarContactosCRMFestivales')
    .addItem('🧠 Depurar contactos con Gemini (seguro)', 'menuDepurarContactosGeminiCRMFestivales')
    .addItem('🛰️ Auditar estructura (seguro)', 'menuAuditarEstructuraCRMFestivales')
    .addItem('🧭 Auditar genero + tamano S/L/XL (seguro)', 'menuAuditarClasificacionCRMFestivales')
    .addSeparator()
    .addItem('⚙️ Instalar trigger de menu', 'instalarTriggerMenuCRMFestivales')
    .addItem('🧹 Limpiar triggers de menu', 'limpiarTriggersMenuCRMFestivales')
    .addSeparator()
    .addItem('📚 Guia de arquitectura', 'mostrarGuiaIntegracionCRMFestivales')
    .addToUi();
}`,
  'menu'
);

replaceOne(
  /function menuHomogeneizarCRMFestivales\(\) \{[\s\S]*?function menuAuditarEstructuraCRMFestivales\(\) \{[\s\S]*?\n\}/,
  `function menuEscanerTotalCRMFestivales() {
  ejecutarConPassword_(homogeneizarCRMFestivales, 'Escaner total + homogeneizar');
}

function menuHomogeneizarCRMFestivales() {
  menuEscanerTotalCRMFestivales();
}

function menuAplicarDisenoCRMFestivales() {
  ejecutarConPassword_(aplicarDisenoCRMFestivales, 'Armonizar diseno visual');
}

function menuDepurarContactosCRMFestivales() {
  ejecutarConPassword_(depurarContactosCRMFestivales, 'Depurar contactos local');
}

function menuDepurarContactosGeminiCRMFestivales() {
  ejecutarConPassword_(depurarContactosConGeminiCRMFestivales, 'Depurar contactos con Gemini');
}

function menuAuditarEstructuraCRMFestivales() {
  ejecutarConPassword_(auditarEstructuraCRMFestivales, 'Auditar estructura');
}

function menuAuditarClasificacionCRMFestivales() {
  ejecutarConPassword_(auditarClasificacionGeneroTamanoCRMFestivales, 'Auditar genero + tamano S/L/XL');
}`,
  'wrappers'
);

replaceOne(
  /function homogeneizarCRMFestivales\(\) \{[\s\S]*?\n\}/,
  `function homogeneizarCRMFestivales() {
  if (!validarSesionSegura_('Escaner total + homogeneizar')) return;
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
    let missingEmail = 0;
    let missingPhone = 0;
    let missingContact = 0;

    sheets.forEach((sheet) => {
      const data = sheet.getDataRange().getValues();
      const normalizedRows = normalizeSheetRows_(data, sheet.getName());
      rewriteSheet_(sheet, normalizedRows);
      applyVisualDesignToSheet_(sheet);
      applyTabVisualIdentity_(sheet);
      const audit = auditarFilasCriticas_(normalizedRows);
      missingEmail += audit.missingEmail;
      missingPhone += audit.missingPhone;
      missingContact += audit.missingContact;
      totalRows += normalizedRows.length;
      totalSheets += 1;
    });

    ordenarPestanasCRMFestivales_(ss);

    SpreadsheetApp.getUi().alert(
      '🚀 Escaner + homogeneizacion completados (' + FEST_ARCHITECT + ').\n\n' +
      'Pestanas procesadas: ' + totalSheets + '\n' +
      'Filas normalizadas: ' + totalRows + '\n' +
      'Sin email: ' + missingEmail + '\n' +
      'Sin telefono: ' + missingPhone + '\n' +
      'Sin contacto: ' + missingContact + '\n\n' +
      'Estructura canonica aplicada en A:J y pestanas armonizadas visualmente.'
    );
  } finally {
    lock.releaseLock();
  }
}`,
  'homogeneizar'
);

replaceOne(
  /function aplicarDisenoCRMFestivales\(\) \{[\s\S]*?\n\}/,
  `function aplicarDisenoCRMFestivales() {
  if (!validarSesionSegura_('Armonizar diseno visual')) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getFestivalSheets_(ss);
  if (!sheets.length) {
    SpreadsheetApp.getUi().alert('No encontre pestanas de festivales para formatear.');
    return;
  }

  sheets.forEach((sheet) => {
    applyVisualDesignToSheet_(sheet);
    applyTabVisualIdentity_(sheet);
  });
  ordenarPestanasCRMFestivales_(ss);

  SpreadsheetApp.getUi().alert('🎨 Diseno visual aplicado a ' + sheets.length + ' pestanas.');
}`,
  'aplicarDiseno'
);

replaceOne(
  /function depurarContactosCRMFestivales\(\) \{[\s\S]*?\n\}/,
  `function depurarContactosCRMFestivales() {
  if (!validarSesionSegura_('Depurar contactos local')) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getFestivalSheets_(ss);
  if (!sheets.length) {
    SpreadsheetApp.getUi().alert('No encontre pestanas de festivales para depurar.');
    return;
  }

  let fixedEmails = 0;
  let fixedPhones = 0;
  let fixedContacts = 0;
  let touchedSheets = 0;

  sheets.forEach((sheet) => {
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return;

    const map = buildHeaderMap_(values[0]);
    const n = values.length - 1;

    const emailOut = [];
    const phoneOut = [];
    const contactOut = [];

    for (let r = 1; r < values.length; r++) {
      const row = values[r];

      const emailRaw = valueAt_(row, map.email);
      const phoneRaw = valueAt_(row, map.telefono);
      const contactRaw = valueAt_(row, map.contacto);

      const email = normalizeEmailCell_(emailRaw);
      const phone = normalizePhoneCell_(phoneRaw);
      const contact = normalizeContactName_(contactRaw);

      if (cleanText_(emailRaw) !== cleanText_(email)) fixedEmails++;
      if (cleanText_(phoneRaw) !== cleanText_(phone)) fixedPhones++;
      if (cleanText_(contactRaw) !== cleanText_(contact)) fixedContacts++;

      emailOut.push([email]);
      phoneOut.push([phone]);
      contactOut.push([contact]);
    }

    sheet.getRange(2, map.email + 1, n, 1).setValues(emailOut);
    sheet.getRange(2, map.telefono + 1, n, 1).setValues(phoneOut);
    sheet.getRange(2, map.contacto + 1, n, 1).setValues(contactOut);

    applyVisualDesignToSheet_(sheet);
    applyTabVisualIdentity_(sheet);
    touchedSheets++;
  });

  ordenarPestanasCRMFestivales_(ss);

  SpreadsheetApp.getUi().alert(
    '✅ Depuracion completada.\n' +
    'Pestanas tocadas: ' + touchedSheets + '\n' +
    'Emails ajustados: ' + fixedEmails + '\n' +
    'Telefonos ajustados: ' + fixedPhones + '\n' +
    'Nombres de contacto ajustados: ' + fixedContacts
  );
}`,
  'depurar local'
);

replaceOne(
  /function auditarEstructuraCRMFestivales\(\) \{[\s\S]*?SpreadsheetApp\.getUi\(\)\.alert\('Auditoria de estructura\\n\\n' \+ lines\.join\('\\n'\)\);\n\}/,
  `function auditarEstructuraCRMFestivales() {
  if (!validarSesionSegura_('Auditar estructura')) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getFestivalSheets_(ss);
  if (!sheets.length) {
    SpreadsheetApp.getUi().alert('No encontre pestanas de festivales para auditar.');
    return;
  }

  const required = FEST_HOMO.HEADER.slice();
  const lines = [];

  sheets.forEach((sheet) => {
    const firstRow = sheet.getRange(1, 1, 1, Math.max(10, sheet.getLastColumn())).getValues()[0];
    const norm = firstRow.map((x) => normalizeHeader_(x));

    const missing = [];
    required.forEach((h) => {
      if (norm.indexOf(normalizeHeader_(h)) === -1) missing.push(h);
    });

    const inOrder = isHeaderInCanonicalOrder_(firstRow);
    if (!missing.length && inOrder) {
      lines.push('✅ ' + sheet.getName() + ': OK');
    } else {
      let msg = '⚠️ ' + sheet.getName() + ': revisar';
      if (missing.length) msg += ' | faltan: ' + missing.join(', ');
      if (!inOrder) msg += ' | orden distinto en A:J';
      lines.push(msg);
    }
  });

  SpreadsheetApp.getUi().alert('🛰️ Auditoria de estructura\\n\\n' + lines.join('\\n'));
}`,
  'auditar estructura'
);

replaceOne(
  /function mostrarGuiaIntegracionCRMFestivales\(\) \{/,
  `function auditarClasificacionGeneroTamanoCRMFestivales() {
  if (!validarSesionSegura_('Auditar genero + tamano S/L/XL')) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getFestivalSheets_(ss);
  if (!sheets.length) {
    SpreadsheetApp.getUi().alert('No encontre pestanas de festivales para auditar clasificacion.');
    return;
  }

  let totalRows = 0;
  let noGenero = 0;
  let noAforo = 0;
  let mismatchGenero = 0;
  let mismatchTamano = 0;
  const examples = [];

  sheets.forEach((sheet) => {
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return;

    const map = buildHeaderMap_(values[0]);
    const tax = parseSheetTaxonomy_(sheet.getName());

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const nombre = cleanText_(valueAt_(row, map.nombre));
      if (!nombre) continue;

      totalRows++;
      const generoRow = normalizeGenreCode_(valueAt_(row, map.genero));
      const sizeRow = sizeCodeFromAforo_(valueAt_(row, map.aforo));

      if (!generoRow) noGenero++;
      if (!sizeRow) noAforo++;

      if (tax.genre && tax.genre !== 'PTE' && generoRow && generoRow !== tax.genre) {
        mismatchGenero++;
        if (examples.length < 18) {
          examples.push(sheet.getName() + '!A' + (r + 1) + ' -> genero=' + generoRow + ' (esperado ' + tax.genre + ')');
        }
      }

      if (tax.size && sizeRow && sizeRow !== tax.size) {
        mismatchTamano++;
        if (examples.length < 18) {
          examples.push(sheet.getName() + '!A' + (r + 1) + ' -> tamano=' + sizeRow + ' (esperado ' + tax.size + ')');
        }
      }
    }
  });

  const resumen = [
    '🧭 Auditoria de clasificacion (' + FEST_ARCHITECT + ')',
    '',
    'Filas revisadas: ' + totalRows,
    'Sin genero: ' + noGenero,
    'Sin aforo/tamano: ' + noAforo,
    'Desajustes de genero: ' + mismatchGenero,
    'Desajustes de tamano: ' + mismatchTamano,
    '',
    'Reglas de tamano:',
    '- S: 0 a 1000 (incluido 1000)',
    '- L: 1001 a 9999',
    '- XL: 10000 en adelante',
    '',
    examples.length ? 'Muestras:\\n' + examples.join('\\n') : 'No se detectaron desajustes en la muestra analizada.'
  ].join('\\n');

  SpreadsheetApp.getUi().alert(resumen);
}

function mostrarGuiaIntegracionCRMFestivales() {`,
  'insert auditoria clasificacion'
);

replaceOne(
  /function getFestivalSheets_\(ss\) \{[\s\S]*?return valid;\n\}/,
  `function getFestivalSheets_(ss) {
  const valid = [];
  const reMain = /^(URBAN|POP|INDIE|ROCK|ELECTR|JAZZ|FLAM|RUMBA|MR|MC|MFR|MEC)_(S|L|XL)$/i;
  const rePending = /^PTE[_\-]/i;

  ss.getSheets().forEach((sheet) => {
    const name = cleanText_(sheet.getName()).toUpperCase();
    if (reMain.test(name) || rePending.test(name)) valid.push(sheet);
  });

  return valid;
}

function parseSheetTaxonomy_(sheetName) {
  const name = cleanText_(sheetName).toUpperCase();
  if (/^PTE[_\-]/.test(name)) return { genre: 'PTE', size: '' };

  const m = name.match(/^(URBAN|POP|INDIE|ROCK|ELECTR|JAZZ|FLAM|RUMBA|MR|MC|MFR|MEC)_(S|L|XL)$/);
  if (!m) return { genre: '', size: '' };

  let genre = m[1];
  if (genre === 'MR') genre = 'MFR';
  if (genre === 'MC') genre = 'MEC';
  return { genre: genre, size: m[2] };
}

function normalizeGenreCode_(raw) {
  const t = normalizeHeader_(raw);
  if (!t) return '';

  if (t.indexOf('URBAN') > -1 || t.indexOf('REGGAE') > -1) return 'URBAN';
  if (t.indexOf('POP') > -1) return 'POP';
  if (t.indexOf('INDIE') > -1) return 'INDIE';
  if (t.indexOf('ROCK') > -1) return 'ROCK';
  if (t.indexOf('ELECTR') > -1) return 'ELECTR';
  if (t.indexOf('JAZZ') > -1) return 'JAZZ';
  if (t.indexOf('FLAM') > -1) return 'FLAM';
  if (t.indexOf('RUMBA') > -1) return 'RUMBA';
  if (t === 'MC' || t === 'MEC' || t.indexOf('CLASICA') > -1 || t.indexOf('CLASICO') > -1) return 'MEC';
  if (t === 'MR' || t === 'MFR' || t.indexOf('REGIONAL') > -1) return 'MFR';
  return '';
}

function normalizeGeneroForDisplay_(raw, fallbackCode) {
  const code = normalizeGenreCode_(raw) || normalizeGenreCode_(fallbackCode);
  if (code) return code;
  return normalizeHeader_(raw);
}

function sizeCodeFromAforo_(aforoRaw) {
  const n = parseAforo_(aforoRaw);
  if (n === '' || isNaN(n)) return '';
  if (n <= 1000) return 'S';
  if (n >= 10000) return 'XL';
  return 'L';
}

function auditarFilasCriticas_(rows) {
  const stats = { missingEmail: 0, missingPhone: 0, missingContact: 0 };
  for (let i = 0; i < rows.length; i++) {
    if (!cleanText_(rows[i][6])) stats.missingEmail++;
    if (!cleanText_(rows[i][7])) stats.missingPhone++;
    if (!cleanText_(rows[i][8])) stats.missingContact++;
  }
  return stats;
}

function applyTabVisualIdentity_(sheet) {
  const tax = parseSheetTaxonomy_(sheet.getName());
  const palette = {
    URBAN: '#FB8C00', POP: '#EC407A', INDIE: '#546E7A', ROCK: '#E53935', ELECTR: '#00ACC1',
    JAZZ: '#3949AB', FLAM: '#D81B60', RUMBA: '#43A047', MEC: '#6D4C41', MFR: '#8D6E63', PTE: '#757575'
  };

  let color = palette[tax.genre] || '#9E9E9E';
  if (tax.size === 'S') color = shiftHex_(color, 35);
  if (tax.size === 'XL') color = shiftHex_(color, -25);
  try { sheet.setTabColor(color); } catch (err) {}
}

function shiftHex_(hex, delta) {
  const h = (hex || '').replace('#', '');
  if (h.length !== 6) return hex;

  const clamp = (v) => Math.max(0, Math.min(255, v));
  const toHex = (v) => clamp(v).toString(16).padStart(2, '0');

  const r = parseInt(h.substring(0, 2), 16);
  const g = parseInt(h.substring(2, 4), 16);
  const b = parseInt(h.substring(4, 6), 16);
  return '#' + toHex(r + delta) + toHex(g + delta) + toHex(b + delta);
}

function legacyGenreCode_(genre) {
  if (genre === 'MEC') return 'MC';
  if (genre === 'MFR') return 'MR';
  return '';
}

function ordenarPestanasCRMFestivales_(ss) {
  const allSheets = ss.getSheets();
  const festSheets = getFestivalSheets_(ss);
  if (!festSheets.length) return;

  const nameToSheet = {};
  festSheets.forEach((sh) => {
    nameToSheet[cleanText_(sh.getName()).toUpperCase()] = sh;
  });

  let firstPos = allSheets.length + 1;
  allSheets.forEach((sh, idx) => {
    if (nameToSheet[cleanText_(sh.getName()).toUpperCase()]) {
      firstPos = Math.min(firstPos, idx + 1);
    }
  });
  if (firstPos > allSheets.length) return;

  const ordered = [];
  FEST_GENRE_ORDER.forEach((genre) => {
    FEST_SIZE_ORDER.forEach((size) => {
      const primary = genre + '_' + size;
      const legacy = legacyGenreCode_(genre) + '_' + size;
      if (nameToSheet[primary]) ordered.push(nameToSheet[primary]);
      else if (legacy !== '_' + size && nameToSheet[legacy]) ordered.push(nameToSheet[legacy]);
    });
  });

  Object.keys(nameToSheet).forEach((name) => {
    if (/^PTE[_\-]/.test(name)) ordered.push(nameToSheet[name]);
  });

  let pos = firstPos;
  ordered.forEach((sh) => {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(pos);
    pos++;
  });
}`,
  'getFestivalSheets + helpers'
);

replaceOne(
  /function normalizeSheetRows_\(data\) \{[\s\S]*?return out;\n\}/,
  `function normalizeSheetRows_(data, sheetName) {
  if (!data || data.length === 0) return [];

  const header = data[0];
  const map = buildHeaderMap_(header);
  const out = [];
  const tax = parseSheetTaxonomy_(sheetName || '');

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const obj = {
      nombre: cleanText_(valueAt_(row, map.nombre)),
      genero: normalizeGeneroForDisplay_(valueAt_(row, map.genero), tax.genre),
      aforo: normalizeAforoForDisplay_(valueAt_(row, map.aforo)),
      ubicacion: cleanText_(valueAt_(row, map.ubicacion)),
      provincia: cleanText_(valueAt_(row, map.provincia)),
      ccaa: cleanText_(valueAt_(row, map.ccaa)),
      email: normalizeEmailCell_(valueAt_(row, map.email)),
      telefono: normalizePhoneCell_(valueAt_(row, map.telefono)),
      contacto: normalizeContactName_(valueAt_(row, map.contacto)),
      observaciones: cleanText_(valueAt_(row, map.observaciones))
    };

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
}`,
  'normalizeSheetRows'
);

replaceOne(
  /function applyVisualDesignToSheet_\(sheet\) \{[\s\S]*?sheet\.setConditionalFormatRules\(\[warnRule, errorRule\]\);\n\}/,
  `function applyVisualDesignToSheet_(sheet) {
  const lastRow = Math.max(2, sheet.getLastRow());
  const lastCol = FEST_HOMO.HEADER.length;

  sheet.setConditionalFormatRules([]);
  sheet.getBandings().forEach((b) => b.remove());

  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange
    .setBackground(FEST_HOMO.COLORS.HEADER_BG)
    .setFontColor(FEST_HOMO.COLORS.HEADER_FG)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, FEST_HOMO.COLORS.GRID, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  const bodyRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  bodyRange
    .setBackground(FEST_HOMO.COLORS.BODY_BG)
    .setFontColor('#222222')
    .setFontFamily('Roboto')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, FEST_HOMO.COLORS.GRID, SpreadsheetApp.BorderStyle.SOLID);

  if (lastRow > 2) {
    const altRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const banding = altRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    banding.setFirstRowColor(FEST_HOMO.COLORS.BODY_BG);
    banding.setSecondRowColor(FEST_HOMO.COLORS.BAND_BG);
  }

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, lastRow, lastCol).createFilter();

  sheet.setColumnWidth(1, 270);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 170);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 260);
  sheet.setColumnWidth(8, 140);
  sheet.setColumnWidth(9, 190);
  sheet.setColumnWidth(10, 260);

  if (lastRow > 1) {
    sheet.getRange(2, 3, lastRow - 1, 1).setHorizontalAlignment('center');
    sheet.getRange(2, 2, lastRow - 1, 1).setHorizontalAlignment('center');
    sheet.getRange(2, 7, lastRow - 1, 2).setHorizontalAlignment('left');

    const generoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(FEST_GENRE_DROPDOWN, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, 2, lastRow - 1, 1).setDataValidation(generoRule);
  }

  const rangeForRules = sheet.getRange(2, 1, Math.max(1, lastRow - 1), lastCol);

  const warnRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"";OR($G2="";$H2=""))')
    .setBackground(FEST_HOMO.COLORS.WARN_BG)
    .setRanges([rangeForRules])
    .build();

  const errorRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"";$G2="";$H2="";$I2="")')
    .setBackground(FEST_HOMO.COLORS.ERROR_BG)
    .setRanges([rangeForRules])
    .build();

  sheet.setConditionalFormatRules([warnRule, errorRule]);
}`,
  'applyVisual'
);

replaceOne(
  /function depurarContactosConGeminiCRMFestivales\(\) \{[\s\S]*?\n\}/,
  `function depurarContactosConGeminiCRMFestivales() {
  if (!validarSesionSegura_('Depurar contactos con Gemini')) return;
  const ui = SpreadsheetApp.getUi();
  const ask = ui.prompt(
    'Depuracion CRM con Gemini',
    'Numero maximo de filas a procesar en esta ejecucion (recomendado 120):',
    ui.ButtonSet.OK_CANCEL
  );
  if (ask.getSelectedButton() !== ui.Button.OK) return;

  const input = parseInt(cleanText_(ask.getResponseText()), 10);
  const maxRows = Number.isFinite(input) && input > 0 ? input : 120;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getFestivalSheets_(ss);
  if (!sheets.length) {
    ui.alert('No encontre pestanas de festivales para depurar con Gemini.');
    return;
  }

  const start = Date.now();
  let reviewed = 0;
  let updated = 0;
  let skipped = 0;
  let errors = 0;
  let lastModelUsed = '';
  const touched = {};

  outer:
  for (let s = 0; s < sheets.length; s++) {
    const sheet = sheets[s];
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) continue;

    const map = buildHeaderMap_(values[0]);

    for (let r = 1; r < values.length; r++) {
      if (reviewed >= maxRows) break outer;
      if (Date.now() - start > FEST_MAX_RUNTIME_MS) break outer;

      const row = values[r];
      const festival = cleanText_(valueAt_(row, map.nombre));
      if (!festival) continue;

      const emailRaw = valueAt_(row, map.email);
      const phoneRaw = valueAt_(row, map.telefono);
      const contactRaw = valueAt_(row, map.contacto);
      const notesRaw = valueAt_(row, map.observaciones);

      const localEmail = normalizeEmailCell_(emailRaw);
      const localPhone = normalizePhoneCell_(phoneRaw);
      const localContact = normalizeContactName_(contactRaw);

      if (!filaNecesitaGemini_(localEmail, localPhone, localContact)) {
        skipped++;
        continue;
      }

      reviewed++;

      const inputObj = {
        nombreFestival: festival,
        genero: cleanText_(valueAt_(row, map.genero)),
        aforo: cleanText_(valueAt_(row, map.aforo)),
        ubicacion: cleanText_(valueAt_(row, map.ubicacion)),
        provincia: cleanText_(valueAt_(row, map.provincia)),
        ccaa: cleanText_(valueAt_(row, map.ccaa)),
        email: localEmail,
        telefono: localPhone,
        nombreContacto: localContact,
        observaciones: cleanText_(notesRaw)
      };

      const aiResult = llamarGeminiDepuracionFila_(inputObj);
      if (!aiResult || !aiResult.data) {
        errors++;
        continue;
      }

      lastModelUsed = aiResult.model || lastModelUsed;
      const ai = aiResult.data;

      let finalEmail = cleanText_(ai.email) ? normalizeEmailCell_(ai.email) : localEmail;
      let finalPhone = cleanText_(ai.telefono) ? normalizePhoneCell_(ai.telefono) : localPhone;
      let finalContact = cleanText_(ai.nombreContacto) ? normalizeContactName_(ai.nombreContacto) : localContact;
      let finalNotes = cleanText_(ai.observaciones) || cleanText_(notesRaw);

      if (!isValidEmailList_(finalEmail)) finalEmail = localEmail;
      if (!formatSpanishPhone_(finalPhone)) finalPhone = localPhone;
      if (!finalContact || isPlaceholderText_(finalContact)) finalContact = localContact;

      const changed =
        cleanText_(finalEmail) !== cleanText_(localEmail) ||
        cleanText_(finalPhone) !== cleanText_(localPhone) ||
        cleanText_(finalContact) !== cleanText_(localContact) ||
        cleanText_(finalNotes) !== cleanText_(notesRaw);

      if (!changed) continue;

      const rowIndex = r + 1;
      sheet.getRange(rowIndex, map.email + 1).setValue(finalEmail);
      sheet.getRange(rowIndex, map.telefono + 1).setValue(finalPhone);
      sheet.getRange(rowIndex, map.contacto + 1).setValue(finalContact);
      if (map.observaciones >= 0) {
        sheet.getRange(rowIndex, map.observaciones + 1).setValue(finalNotes);
      }

      touched[sheet.getName()] = true;
      updated++;
    }
  }

  Object.keys(touched).forEach((name) => {
    const sh = ss.getSheetByName(name);
    if (sh) {
      applyVisualDesignToSheet_(sh);
      applyTabVisualIdentity_(sh);
    }
  });

  ordenarPestanasCRMFestivales_(ss);

  const timeoutReached = Date.now() - start > FEST_MAX_RUNTIME_MS;
  ui.alert(
    '🧠 Depuracion con Gemini finalizada.\\n\\n' +
    'Filas revisadas por IA: ' + reviewed + '\\n' +
    'Filas actualizadas: ' + updated + '\\n' +
    'Filas sin cambios: ' + skipped + '\\n' +
    'Errores IA: ' + errors + '\\n' +
    'Modelo principal configurado: gemini-3.1-pro-preview\\n' +
    'Ultimo modelo usado: ' + (lastModelUsed || 'No disponible') + '\\n' +
    (timeoutReached ? '\\nSe alcanzo el tiempo maximo de Apps Script. Ejecuta de nuevo para continuar.' : '')
  );
}`,
  'depurar gemini'
);

replaceOne(
  /function normalizeContactName_\(v\) \{[\s\S]*?return out\.join\(' '\);\n\}/,
  `function normalizeContactName_(v) {
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

function isValidEmailList_(raw) {
  const txt = cleanText_(raw).toLowerCase();
  if (!txt) return false;
  const tokens = txt.split(/[;,|\s]+/).map((x) => x.trim()).filter((x) => x);
  if (!tokens.length) return false;
  for (let i = 0; i < tokens.length; i++) {
    if (!isValidEmail_(tokens[i])) return false;
  }
  return true;
}

function isPlaceholderText_(v) {
  const t = cleanText_(v).toLowerCase();
  if (!t) return true;
  return /^(sin informacion|sin info|n\/a|na|s\/d|desconocido|none|-+)$/i.test(t);
}

function filaNecesitaGemini_(email, phone, contact) {
  const emailOk = isValidEmailList_(email);
  const phoneOk = !!formatSpanishPhone_(phone);
  const contactOk = !!cleanText_(contact) && !isPlaceholderText_(contact);
  return !emailOk || !phoneOk || !contactOk;
}`,
  'helpers contacto'
);

replaceOne(
  /<h2 style="margin-top:0;">Como conectar esta hoja con codigo<\/h2>/,
  `<h2 style="margin-top:0;">🚀 Arquitectura CRM FESTIVALES</h2><p><b>ARQUITECTO:</b> ` + `' + FEST_ARCHITECT + '` + `</p>`,
  'guia titulo'
);

replaceOne(
  /Menu CRM Festivales -> homogeneizar -> depurar contactos -> auditar estructura\./,
  'Menu 🚀 CRM FESTIVALES -> escaner total -> depurar contactos -> auditorias de estructura y clasificacion.',
  'guia flujo'
);

fs.writeFileSync(path, t, 'utf8');
console.log('OK patch');
