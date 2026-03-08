const fs = require('fs');
const path = 'C:/Users/elrub/Desktop/CARPETA CODEX/01_PROYECTOS/festivales-github/Código.js';
let text = fs.readFileSync(path, 'utf8').replace(/\r\n/g, '\n');

function replaceOrThrow(regex, replacement, label) {
  const prev = text;
  text = text.replace(regex, replacement);
  if (prev === text) {
    throw new Error('No se encontro bloque para: ' + label);
  }
}

replaceOrThrow(
  /const FEST_MAX_RUNTIME_MS = 4\.7 \* 60 \* 1000;\n/,
`const FEST_MAX_RUNTIME_MS = 4.7 * 60 * 1000;
const FEST_ARCHITECT = 'RUBEN COTON';
const FEST_GENRE_ORDER = ['URBAN', 'POP', 'INDIE', 'ROCK', 'ELECTR', 'JAZZ', 'FLAM', 'RUMBA', 'MEC', 'MFR'];
const FEST_SIZE_ORDER = ['S', 'L', 'XL'];
const FEST_GENRE_DROPDOWN = [
  '🧢 URBAN',
  '🎤 POP',
  '🎸 INDIE',
  '🤘 ROCK',
  '🎛️ ELECTR',
  '🎷 JAZZ',
  '💃 FLAM',
  '🪘 RUMBA',
  '🎼 MEC',
  '🌄 MFR'
];
`,
  'constantes'
);

replaceOrThrow(
  /function crearMenuCRMFestivales_\(\) \{[\s\S]*?\n\}\n\nfunction instalarTriggerMenuCRMFestivales\(\) \{/,
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
}

function instalarTriggerMenuCRMFestivales() {`,
  'menu'
);

replaceOrThrow(
  /A partir de ahora, al abrir la hoja, aparecera el menu CRM Festivales\./,
  'A partir de ahora, al abrir la hoja, aparecera el menu 🚀 CRM FESTIVALES | RUBEN COTON.',
  'mensaje trigger'
);

replaceOrThrow(
  /function menuHomogeneizarCRMFestivales\(\) \{[\s\S]*?\n\}\n\n\/\*\*[\s\S]*?\* 1\)/,
`function menuHomogeneizarCRMFestivales() {
  menuEscanerTotalCRMFestivales();
}

function menuEscanerTotalCRMFestivales() {
  ejecutarConPassword_(homogeneizarCRMFestivales, 'Escaner total + homogeneizar');
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
}

/**
 * 1)`,
  'wrappers menu'
);

replaceOrThrow(
  /function homogeneizarCRMFestivales\(\) \{[\s\S]*?\n\}\n\n\/\*\*\n \* Reaplica solo el formato visual a las pestanas detectadas\./,
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

      const stats = auditarFilasCriticas_(normalizedRows);
      missingEmail += stats.missingEmail;
      missingPhone += stats.missingPhone;
      missingContact += stats.missingContact;

      totalRows += normalizedRows.length;
      totalSheets += 1;
    });

    ordenarPestanasCRMFestivales_(ss);

    const resumenEscaner = [
      '🚀 Escaner + homogeneizacion completados (' + FEST_ARCHITECT + ').',
      '',
      'Pestanas procesadas: ' + totalSheets,
      'Filas normalizadas: ' + totalRows,
      'Sin email: ' + missingEmail,
      'Sin telefono: ' + missingPhone,
      'Sin contacto: ' + missingContact,
      '',
      'Estructura canonica aplicada en A:J y pestanas armonizadas visualmente.'
    ].join('\n');

    SpreadsheetApp.getUi().alert(resumenEscaner);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Reaplica solo el formato visual a las pestanas detectadas.`,
  'homogeneizar'
);

replaceOrThrow(
  /function aplicarDisenoCRMFestivales\(\) \{[\s\S]*?\n\}\n\n\/\*\*\n \* Solo depura campos de contacto sin reordenar filas\./,
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
}

/**
 * Solo depura campos de contacto sin reordenar filas.`,
  'aplicar diseno'
);

replaceOrThrow(
  /function depurarContactosCRMFestivales\(\) \{[\s\S]*?\n\}\n\nfunction auditarEstructuraCRMFestivales\(\) \{/,
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
}

function auditarEstructuraCRMFestivales() {`,
  'depurar contactos local'
);

replaceOrThrow(
  /function auditarEstructuraCRMFestivales\(\) \{[\s\S]*?\n\}\n\nfunction mostrarGuiaIntegracionCRMFestivales\(\) \{/,
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

  SpreadsheetApp.getUi().alert('🛰️ Auditoria de estructura\n\n' + lines.join('\n'));
}

function auditarClasificacionGeneroTamanoCRMFestivales() {
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
    examples.length ? 'Muestras:\n' + examples.join('\n') : 'No se detectaron desajustes en la muestra analizada.'
  ].join('\n');

  SpreadsheetApp.getUi().alert(resumen);
}

function mostrarGuiaIntegracionCRMFestivales() {`,
  'auditorias'
);

replaceOrThrow(
  /function mostrarGuiaIntegracionCRMFestivales\(\) \{[\s\S]*?\n\}\n\nfunction getFestivalSheets_\(ss\) \{/,
`function mostrarGuiaIntegracionCRMFestivales() {
  const html = [
    '<div style="font-family:Arial,sans-serif;padding:14px;line-height:1.5;color:#222;">',
    '<h2 style="margin-top:0;">🚀 Arquitectura CRM FESTIVALES</h2><p><b>ARQUITECTO:</b> ' + FEST_ARCHITECT + '</p>',
    '<p><b>1) En la propia hoja (Apps Script)</b><br>Extensiones -> Apps Script. Editas funciones y ejecutas desde menu o triggers.</p>',
    '<p><b>2) Con repositorio local (clasp)</b><br>Puedes sincronizar el proyecto de Apps Script con archivos .gs en tu ordenador y versionarlo con Git.</p>',
    '<p><b>3) Con APIs externas</b><br>Tu script puede llamar APIs (Gemini u otras) con UrlFetchApp y guardar resultados en celdas.</p>',
    '<p><b>4) Modificaciones seguras recomendadas</b><br>Crea copia de la hoja antes de cambios grandes, usa una pestana de pruebas, y luego aplicas a produccion.</p>',
    '<p><b>5) Flujo recomendado para ti</b><br>Menu 🚀 CRM FESTIVALES -> escaner total -> depurar contactos -> auditorias de estructura y clasificacion.</p>',
    '</div>'
  ].join('');

  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(html).setWidth(520).setHeight(380),
    'Guia CRM Festivales'
  );
}

function getFestivalSheets_(ss) {`,
  'guia'
);

replaceOrThrow(
  /function getFestivalSheets_\(ss\) \{[\s\S]*?\n\}\n\nfunction normalizeSheetRows_\(data\) \{/,
`function getFestivalSheets_(ss) {
  const valid = [];
  const reMain = /^(URBAN|POP|INDIE|ROCK|ELECTR|JAZZ|FLAM|RUMBA|MR|MC|MFR|MEC)_(S|L|XL)$/i;
  const rePending = /^PTE[_-]/i;

  ss.getSheets().forEach((sheet) => {
    const name = cleanText_(sheet.getName()).toUpperCase();
    if (reMain.test(name) || rePending.test(name)) {
      valid.push(sheet);
    }
  });

  return valid;
}

function parseSheetTaxonomy_(sheetName) {
  const name = cleanText_(sheetName).toUpperCase();
  if (/^PTE[_-]/.test(name)) return { genre: 'PTE', size: '' };

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
    if (/^PTE[_-]/.test(name)) ordered.push(nameToSheet[name]);
  });

  let pos = firstPos;
  ordered.forEach((sh) => {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(pos);
    pos++;
  });
}

function normalizeSheetRows_(data, sheetName) {`,
  'taxonomy + orden'
);

replaceOrThrow(
  /function normalizeSheetRows_\(data, sheetName\) \{[\s\S]*?\n\}\n\nfunction rewriteSheet_\(sheet, rows\) \{/,
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
}

function rewriteSheet_(sheet, rows) {`,
  'normalize rows'
);

replaceOrThrow(
  /function applyVisualDesignToSheet_\(sheet\) \{[\s\S]*?\n\}\n\nfunction buildHeaderMap_\(headerRow\) \{/,
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
}

function buildHeaderMap_(headerRow) {`,
  'apply visual design'
);

replaceOrThrow(
  /Object\.keys\(touched\)\.forEach\(\(name\) => \{\n\s*const sh = ss\.getSheetByName\(name\);\n\s*if \(sh\) applyVisualDesignToSheet_\(sh\);\n\s*\}\);/,
`Object.keys(touched).forEach((name) => {
    const sh = ss.getSheetByName(name);
    if (sh) {
      applyVisualDesignToSheet_(sh);
      applyTabVisualIdentity_(sh);
    }
  });

  ordenarPestanasCRMFestivales_(ss);`,
  'post Gemini visual'
);

fs.writeFileSync(path, text, 'utf8');
console.log('OK');
