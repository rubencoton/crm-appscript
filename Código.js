// =============================================================================
// CRM FESTIVALES - MENU OPERATIVO + HOMOGENEIDAD + DISENO
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
    GRID: '#D9D9D9'
  }
};

const FEST_SECURITY_PASSWORD = '+rubencoton26';
const FEST_GEMINI_API_KEY = 'AIzaSyC2AnnQuFgKOR_qGNl4jTrsoWF672bnK0M';
const FEST_GEMINI_MODELS = [
  'gemini-3.1-pro-preview',
  'gemini-2.5-pro',
  'gemini-2.5-flash',
  'gemini-1.5-pro-latest'
];
const FEST_MAX_RUNTIME_MS = 4.7 * 60 * 1000;

// Si no tienes otro onOpen en el proyecto, este ya deja el menu automatico.

function onOpen(e) {
  crearMenuCRMFestivales_();
}

// Alias para compatibilidad con versiones previas del script.
function onOpenFestivalesHomogeneidad_() {
  crearMenuCRMFestivales_();
}

function crearMenuCRMFestivales_() {
  SpreadsheetApp.getUi()
    .createMenu('CRM Festivales')
    .addItem('1) Homogeneizar columnas + diseno (seguro)', 'menuHomogeneizarCRMFestivales')
    .addItem('2) Solo aplicar diseno visual (seguro)', 'menuAplicarDisenoCRMFestivales')
    .addItem('3) Depurar contactos local (seguro)', 'menuDepurarContactosCRMFestivales')
    .addItem('4) Depurar contactos con Gemini (seguro)', 'menuDepurarContactosGeminiCRMFestivales')
    .addItem('5) Auditar estructura de pestanas (seguro)', 'menuAuditarEstructuraCRMFestivales')
    .addSeparator()
    .addItem('6) Instalar trigger de menu', 'instalarTriggerMenuCRMFestivales')
    .addItem('7) Limpiar triggers de menu', 'limpiarTriggersMenuCRMFestivales')
    .addSeparator()
    .addItem('8) Guia: conectar con mas codigo', 'mostrarGuiaIntegracionCRMFestivales')
    .addToUi();
}

function instalarTriggerMenuCRMFestivales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ScriptApp.getProjectTriggers();
  let deleted = 0;

  existing.forEach((tr) => {
    const fn = tr.getHandlerFunction();
    if (fn === 'crearMenuCRMFestivales_') {
      ScriptApp.deleteTrigger(tr);
      deleted++;
    }
  });

  ScriptApp.newTrigger('crearMenuCRMFestivales_').forSpreadsheet(ss).onOpen().create();

  SpreadsheetApp.getUi().alert(
    'Trigger de menu instalado correctamente.\n' +
    'Triggers antiguos eliminados: ' + deleted + '\n\n' +
    'A partir de ahora, al abrir la hoja, aparecera el menu CRM Festivales.'
  );
}

function limpiarTriggersMenuCRMFestivales() {
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = 0;

  triggers.forEach((tr) => {
    const fn = tr.getHandlerFunction();
    if (fn === 'crearMenuCRMFestivales_') {
      ScriptApp.deleteTrigger(tr);
      deleted++;
    }
  });

  SpreadsheetApp.getUi().alert('Triggers de menu eliminados: ' + deleted);
}


function ejecutarConPassword_(accionFn, etiqueta) {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt(
    'Seguridad CRM Festivales',
    'Introduce la contrasena para: ' + etiqueta,
    ui.ButtonSet.OK_CANCEL
  );

  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  const pass = cleanText_(prompt.getResponseText());
  if (pass !== FEST_SECURITY_PASSWORD) {
    ui.alert('Contrasena incorrecta. Accion cancelada.');
    return;
  }

  activarSesionSegura_();
  accionFn();
}

function activarSesionSegura_() {
  try {
    CacheService.getUserCache().put('FEST_AUTH_OK', '1', 300);
  } catch (err) {
    // no-op
  }
}

function validarSesionSegura_(accion) {
  try {
    if (CacheService.getUserCache().get('FEST_AUTH_OK') === '1') return true;
  } catch (err) {
    // no-op
  }

  SpreadsheetApp.getUi().alert(
    'Acceso denegado para: ' + accion + '\n\nPrimero ejecuta la accion desde el menu seguro e introduce la contrasena.'
  );
  return false;
}

function menuHomogeneizarCRMFestivales() {
  ejecutarConPassword_(homogeneizarCRMFestivales, 'Homogeneizar columnas + diseno');
}

function menuAplicarDisenoCRMFestivales() {
  ejecutarConPassword_(aplicarDisenoCRMFestivales, 'Aplicar diseno visual');
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

/**
 * 1) Unifica TODAS las pestanas de festivales al mismo orden de columnas.
 * 2) Aplica diseno visual consistente para una lectura comoda.
 */
function homogeneizarCRMFestivales() {
  if (!validarSesionSegura_('Homogeneizar columnas + diseno')) return;
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
  if (!validarSesionSegura_('Aplicar diseno visual')) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getFestivalSheets_(ss);
  if (!sheets.length) {
    SpreadsheetApp.getUi().alert('No encontre pestanas de festivales para formatear.');
    return;
  }

  sheets.forEach((sheet) => applyVisualDesignToSheet_(sheet));
  SpreadsheetApp.getUi().alert('Diseno visual aplicado a ' + sheets.length + ' pestanas.');
}

/**
 * Solo depura campos de contacto sin reordenar filas.
 */
function depurarContactosCRMFestivales() {
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
    const range = sheet.getDataRange();
    const values = range.getValues();
    if (values.length < 2) return;

    const header = values[0];
    const map = buildHeaderMap_(header);
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
    touchedSheets++;
  });

  SpreadsheetApp.getUi().alert(
    'Depuracion completada.\n' +
    'Pestanas tocadas: ' + touchedSheets + '\n' +
    'Emails ajustados: ' + fixedEmails + '\n' +
    'Telefonos ajustados: ' + fixedPhones + '\n' +
    'Nombres de contacto ajustados: ' + fixedContacts
  );
}

function auditarEstructuraCRMFestivales() {
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
      lines.push(sheet.getName() + ': OK');
    } else {
      let msg = sheet.getName() + ': revisar';
      if (missing.length) msg += ' | faltan: ' + missing.join(', ');
      if (!inOrder) msg += ' | orden distinto en A:J';
      lines.push(msg);
    }
  });

  SpreadsheetApp.getUi().alert('Auditoria de estructura\n\n' + lines.join('\n'));
}

function mostrarGuiaIntegracionCRMFestivales() {
  const html = [
    '<div style="font-family:Arial,sans-serif;padding:14px;line-height:1.5;color:#222;">',
    '<h2 style="margin-top:0;">Como conectar esta hoja con codigo</h2>',
    '<p><b>1) En la propia hoja (Apps Script)</b><br>Extensiones -> Apps Script. Editas funciones y ejecutas desde menu o triggers.</p>',
    '<p><b>2) Con repositorio local (clasp)</b><br>Puedes sincronizar el proyecto de Apps Script con archivos .gs en tu ordenador y versionarlo con Git.</p>',
    '<p><b>3) Con APIs externas</b><br>Tu script puede llamar APIs (Gemini u otras) con UrlFetchApp y guardar resultados en celdas.</p>',
    '<p><b>4) Modificaciones seguras recomendadas</b><br>Crea copia de la hoja antes de cambios grandes, usa una pestana de pruebas, y luego aplicas a produccion.</p>',
    '<p><b>5) Flujo recomendado para ti</b><br>Menu CRM Festivales -> homogeneizar -> depurar contactos -> auditar estructura.</p>',
    '</div>'
  ].join('');

  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(html).setWidth(520).setHeight(380),
    'Guia CRM Festivales'
  );
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
  sheet.clear();
  sheet.getRange(1, 1, 1, FEST_HOMO.HEADER.length).setValues([FEST_HOMO.HEADER]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, FEST_HOMO.HEADER.length).setValues(rows);
  }
}

function applyVisualDesignToSheet_(sheet) {
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
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, lastRow, lastCol).createFilter();

  sheet.setColumnWidth(1, 270);
  sheet.setColumnWidth(2, 110);
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

function isHeaderInCanonicalOrder_(headerRow) {
  const expected = FEST_HOMO.HEADER.map((x) => normalizeHeader_(x));
  for (let i = 0; i < expected.length; i++) {
    const cell = i < headerRow.length ? normalizeHeader_(headerRow[i]) : '';
    if (cell !== expected[i]) return false;
  }
  return true;
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




function depurarContactosConGeminiCRMFestivales() {
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
    if (sh) applyVisualDesignToSheet_(sh);
  });

  const timeoutReached = Date.now() - start > FEST_MAX_RUNTIME_MS;
  ui.alert(
    'Depuracion con Gemini finalizada.\n\n' +
    'Filas revisadas por IA: ' + reviewed + '\n' +
    'Filas actualizadas: ' + updated + '\n' +
    'Filas sin cambios: ' + skipped + '\n' +
    'Errores IA: ' + errors + '\n' +
    'Modelo principal configurado: gemini-3.1-pro-preview\n' +
    'Ultimo modelo usado: ' + (lastModelUsed || 'No disponible') + '\n' +
    (timeoutReached ? '\nSe alcanzo el tiempo maximo de Apps Script. Ejecuta de nuevo para continuar.' : '')
  );
}

function filaNecesitaGemini_(email, phone, contact) {
  const emailOk = isValidEmailList_(email);
  const phoneOk = !!formatSpanishPhone_(phone);
  const contactOk = !!cleanText_(contact) && !isPlaceholderText_(contact);
  return !emailOk || !phoneOk || !contactOk;
}

function isPlaceholderText_(v) {
  const t = cleanText_(v).toLowerCase();
  if (!t) return true;
  return /^(sin informacion|sin info|n\/a|na|s\/d|desconocido|none|-+)$/i.test(t);
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

function llamarGeminiDepuracionFila_(rowObj) {
  const schema = {
    type: 'OBJECT',
    properties: {
      email: { type: 'STRING' },
      telefono: { type: 'STRING' },
      nombreContacto: { type: 'STRING' },
      observaciones: { type: 'STRING' }
    },
    required: ['email', 'telefono', 'nombreContacto', 'observaciones']
  };

  const prompt = [
    'Depura SOLO estos campos de CRM: email, telefono, nombreContacto, observaciones.',
    'Reglas estrictas:',
    '1) No inventes datos. Si no hay evidencia suficiente, devuelve el valor de entrada sin cambios.',
    '2) Email: deja uno o varios separados por "; ", todos validos.',
    '3) Telefono: formato espana +34 XXX XXX XXX si es posible; si no, conserva original.',
    '4) nombreContacto: capitalizacion correcta, sin texto basura.',
    '5) Devuelve JSON puro con las 4 claves requeridas.',
    '',
    'Entrada JSON:',
    JSON.stringify(rowObj)
  ].join('\n');

  return invocarGeminiConFallback_(prompt, schema);
}

function invocarGeminiConFallback_(prompt, responseSchema) {
  const payload = {
    systemInstruction: {
      parts: [{
        text: [
          'Eres un analista de calidad de datos CRM en espanol.',
          'Nunca inventes datos no presentes en la entrada.',
          'Responde SOLO en JSON valido, sin markdown.'
        ].join(' ')
      }]
    },
    contents: [
      {
        role: 'user',
        parts: [{ text: prompt }]
      }
    ],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 800,
      responseMimeType: 'application/json',
      responseSchema: responseSchema
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  for (let i = 0; i < FEST_GEMINI_MODELS.length; i++) {
    const model = FEST_GEMINI_MODELS[i];
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
      model + ':generateContent?key=' + FEST_GEMINI_API_KEY;

    try {
      const res = UrlFetchApp.fetch(url, options);
      const code = res.getResponseCode();
      const raw = res.getContentText();

      if (code === 200) {
        const parsed = parseGeminiJson_(raw);
        if (parsed) {
          return { model: model, data: parsed };
        }
      }

      if (code === 404 || code === 429 || code >= 500) {
        Utilities.sleep(900);
        continue;
      }
    } catch (err) {
      Utilities.sleep(900);
    }
  }

  return null;
}

function parseGeminiJson_(rawText) {
  try {
    const root = JSON.parse(rawText);
    if (!root || !root.candidates || !root.candidates.length) return null;
    const part = root.candidates[0].content && root.candidates[0].content.parts
      ? root.candidates[0].content.parts[0]
      : null;
    if (!part) return null;

    if (part.text) {
      const match = part.text.match(/\{[\s\S]*\}/);
      if (!match) return null;
      return JSON.parse(match[0]);
    }

    return null;
  } catch (err) {
    return null;
  }
}



