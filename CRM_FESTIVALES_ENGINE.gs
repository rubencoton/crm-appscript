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
    'OBSERVACIONES',
    'REVISION EMAIL',
    'Merge status'
  ],
  COLORS: {
    HEADER_BG: '#8B0000',
    HEADER_FG: '#FFFFFF',
    BODY_BG: '#FFFFFF',
    BAND_BG: '#FFF9CC',
    WARN_BG: '#FFF4CC',
    ERROR_BG: '#FDE2E2',
    GRID: '#D9D9D9'
  }
};

// IMPORTANTE: mantener vacio en repositorio publico.
const FEST_SECURITY_PASSWORD = '';
const FEST_GEMINI_API_KEY = '';
const FEST_GEMINI_MODELS = [
  'gemini-3.1-pro-preview',
  'gemini-2.5-pro',
  'gemini-3-flash-preview',
  'gemini-2.5-flash',
  'gemini-2.5-flash-lite',
  'gemini-flash-latest'
];
const FEST_GEMINI_MAX_RETRIES_PER_MODEL = 2;
const FEST_GEMINI_BASE_MAX_OUTPUT_TOKENS = 2048;
const FEST_GEMINI_RETRY_MAX_OUTPUT_TOKENS = 4096;
const FEST_GEMINI_RESPONSE_CACHE_TTL_SEC = 6 * 60 * 60;
const FEST_MAX_RUNTIME_MS = 4.7 * 60 * 1000;
const FEST_ARCHITECT = 'RUBEN COTON';
const FEST_MERGE_STATUS_HEADER = 'Merge status';
const FEST_GENRE_DROPDOWN = [
  'URBANO', 'URBAN',
  'POP', 'INDIE', 'ROCK',
  'ELECTRONICA', 'ELECTR',
  'JAZZ',
  'FLAMENCO', 'FLAM',
  'RUMBA',
  'MFR', 'MR',
  'MEC', 'MC',
  'PENDIENTE', 'PTE'
];
const FEST_EMAIL_REVIEW_OPTIONS = [
  'BIEN',
  'CORREGIDO',
  'MAL'
];
const FEST_EMAIL_REVIEW_SHEET_NAME = 'REVISION_EMAILS';
const FEST_EMAIL_TRIGGER_HANDLER = 'auditarEmailsAutomaticaCRMFestivales_';
const FEST_EMAIL_OPEN_TRIGGER_HANDLER = 'auditarEmailsAlAbrirCRMFestivales_';
const FEST_EMAIL_TRIGGER_INTERVAL_HOURS = 6;
const FEST_EMAIL_OPEN_COOLDOWN_MINUTES = 2;
const FEST_EMAIL_RUN_COOLDOWN_MINUTES = 2;
const FEST_EMAIL_MAX_DOMAIN_CHECKS_PER_RUN = 140;
const FEST_EMAIL_DOMAIN_CACHE_TTL_HOURS = 24;
const FEST_EMAIL_MAX_WEB_FETCHES_PER_RUN = 40;
const FEST_EMAIL_MAX_SEARCH_FETCHES_PER_RUN = 10;
const FEST_EMAIL_WEB_CACHE_TTL_SEC = 8 * 60 * 60;

// Si no tienes otro onOpen en el proyecto, este ya deja el menu automatico.

/*
 Credenciales y seguridad de este motor:
 - FEST_SECURITY_PASSWORD y FEST_GEMINI_API_KEY se guardan en Script Properties.
 - El menu "Configurar credenciales Gemini" permite cargarlas sin exponerlas en codigo.
 - Las constantes locales deben permanecer vacias en repositorio publico.
*/
function getFestScriptProperty_(key) {
  try {
    return cleanText_(PropertiesService.getScriptProperties().getProperty(key));
  } catch (err) {
    return '';
  }
}

function getFestSecurityPassword_() {
  const pass = getFestScriptProperty_('FEST_SECURITY_PASSWORD') || cleanText_(FEST_SECURITY_PASSWORD);
  if (!pass) {
    throw new Error('Falta FEST_SECURITY_PASSWORD. Usa el menu: "Configurar credenciales Gemini".');
  }
  return pass;
}

function getFestGeminiApiKey_() {
  return getFestScriptProperty_('FEST_GEMINI_API_KEY') || cleanText_(FEST_GEMINI_API_KEY);
}

function onOpen(e) {
  crearMenuCRMFestivales_();
  try {
    aplicarAjustesVisualesAlAbrirCRMFestivales_();
  } catch (err) {
    Logger.log('onOpen visual refresh skip: ' + (err && err.message ? err.message : err));
  }
  try {
    desactivarTriggersAutoRevisionEmailsSilencioso_();
  } catch (err) {
    Logger.log('onOpen auto-trigger cleanup skip: ' + (err && err.message ? err.message : err));
  }
}

// Alias para compatibilidad con versiones previas del script.
function onOpenFestivalesHomogeneidad_() {
  crearMenuCRMFestivales_();
}

function crearMenuCRMFestivales_() {
  SpreadsheetApp.getUi()
    .createMenu('CRM FESTIVALES | RUBEN COTON')
    .addItem('BOTON | Auditar contactos web (bloquea + progreso)', 'botonAuditarContactosWebCRMFestivales')
    .addItem('BOTON | Autocompletado de celdas (IA)', 'botonAutocompletadoCeldasIA_CRMFestivales')
    .addToUi();
}

function aplicarAjustesVisualesAlAbrirCRMFestivales_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(500)) return;
  try {
    const started = Date.now();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    const sheets = getFestivalSheets_(ss);
    if (!sheets.length) {
      const active = ss.getActiveSheet();
      if (active && isFestivalSheetName_(active.getName())) sheets.push(active);
    }
    if (!sheets.length) return;
    const activeSheet = ss.getActiveSheet();
    const activeSheetId = activeSheet ? activeSheet.getSheetId() : null;
    const order = ordenarSheetsParaAjusteRapido_(sheets, activeSheetId);
    const orderedSheets = order.list;
    const hasActivePinned = order.activePinned;

    const hardLimitMs = 4500;
    let applied = 0;
    let skipped = 0;
    let advancedCursor = 0;
    for (let i = 0; i < orderedSheets.length; i++) {
      if (Date.now() - started > hardLimitMs) {
        skipped += Math.max(0, orderedSheets.length - i);
        break;
      }
      const isPinnedActive = hasActivePinned && i === 0;
      if (!isPinnedActive) advancedCursor++;
      try {
        if (aplicarAjustesVisualesRapidosEnSheet_(orderedSheets[i])) applied++;
        else skipped++;
      } catch (err) {
        skipped++;
      }
    }
    actualizarCursorAjusteRapido_(order.cursorStart, advancedCursor, order.rotatableCount);
    Logger.log('Ajuste visual rapido onOpen ms=' + (Date.now() - started) + ' | aplicadas=' + applied + ' | omitidas=' + skipped);
  } finally {
    lock.releaseLock();
  }
}

function ordenarSheetsParaAjusteRapido_(sheets, activeSheetId) {
  const list = Array.isArray(sheets) ? sheets.slice() : [];
  let activePinned = false;
  let activeSheet = null;
  if (activeSheetId) {
    const idx = list.findIndex((x) => x && x.getSheetId && x.getSheetId() === activeSheetId);
    if (idx >= 0) {
      activePinned = true;
      activeSheet = list[idx];
      list.splice(idx, 1);
    }
  }

  const rotatableCount = list.length;
  let cursorStart = 0;
  if (rotatableCount > 0) {
    try {
      const raw = Number(PropertiesService.getScriptProperties().getProperty('FEST_FAST_OPEN_CURSOR') || '0');
      cursorStart = ((raw % rotatableCount) + rotatableCount) % rotatableCount;
    } catch (err) {
      cursorStart = 0;
    }
  }

  const rotated = rotatableCount > 0
    ? list.slice(cursorStart).concat(list.slice(0, cursorStart))
    : [];

  const out = activePinned && activeSheet ? [activeSheet].concat(rotated) : rotated;
  return {
    list: out,
    activePinned: activePinned && !!activeSheet,
    cursorStart: cursorStart,
    rotatableCount: rotatableCount
  };
}

function actualizarCursorAjusteRapido_(cursorStart, advanced, rotatableCount) {
  if (!rotatableCount || rotatableCount <= 0) return;
  const step = Math.max(1, Number(advanced) || 0);
  const next = (Math.max(0, Number(cursorStart) || 0) + step) % rotatableCount;
  try {
    PropertiesService.getScriptProperties().setProperty('FEST_FAST_OPEN_CURSOR', String(next));
  } catch (err) {
    // no-op
  }
}

function aplicarAjustesVisualesRapidosEnSheet_(sh) {
  if (!sh || !isFestivalSheetName_(sh.getName())) return false;
  const initialCols = Math.max(sh.getLastColumn(), FEST_HOMO.HEADER.length);
  const initialHeader = sh.getRange(1, 1, 1, initialCols).getValues()[0];
  if (!isHeaderInCanonicalOrder_(initialHeader)) {
    if (!repararCabecerasMinimasEnSheet_(sh, initialHeader)) return false;
  }

  const lastCol = Math.max(sh.getLastColumn(), FEST_HOMO.HEADER.length);
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  const map = buildHeaderMap_(header);
  const reviewCol = map.reviewEmail + 1;
  const mergeCol = map.mergeStatus + 1;
  const genreCol = map.genero + 1;
  const usedCols = Math.max(lastCol, reviewCol, mergeCol);

  sh.getRange(1, 1, 1, usedCols)
    .setBackground(FEST_HOMO.COLORS.HEADER_BG)
    .setFontColor(FEST_HOMO.COLORS.HEADER_FG)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sh.getRange(1, mergeCol, 1, 1).setValue(FEST_MERGE_STATUS_HEADER);
  sh.getRange(1, reviewCol, 1, 1)
    .setValue('REVISION EMAIL')
    .setBackground('#FBC02D')
    .setFontColor('#8B0000')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const rows = Math.max(0, sh.getLastRow() - 1);
  if (rows > 0) {
    const generoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(FEST_GENRE_DROPDOWN, true)
      .setAllowInvalid(false)
      .build();
    sh.getRange(2, genreCol, rows, 1)
      .setDataValidation(generoRule)
      .setHorizontalAlignment('center');
    sh.getRange(2, reviewCol, rows, 1)
      .setDataValidation(buildEmailReviewValidationRule_())
      .setHorizontalAlignment('center')
      .setFontWeight('bold');
    sh.getRange(2, mergeCol, rows, 1)
      .clearDataValidations()
      .setHorizontalAlignment('left')
      .setFontWeight('normal');
  }

  return true;
}

function repararCabecerasMinimasEnSheet_(sheet, headerRow) {
  if (!sheet) return false;
  let cols = Math.max(sheet.getLastColumn(), FEST_HOMO.HEADER.length);
  let header = Array.isArray(headerRow) ? headerRow.slice() : sheet.getRange(1, 1, 1, cols).getValues()[0];

  function findHeaderIndex(names) {
    const wanted = {};
    for (let i = 0; i < names.length; i++) wanted[normalizeHeader_(names[i])] = true;
    for (let c = 0; c < header.length; c++) {
      const h = normalizeHeader_(header[c]);
      if (wanted[h]) return c + 1;
    }
    return 0;
  }

  let reviewCol = findHeaderIndex(['REVISION EMAIL', 'ESTADO EMAIL', 'EMAIL REVISADO']);
  let mergeCol = findHeaderIndex(['MERGE STATUS', 'MERGESTATUS', 'MERGE']);

  if (!reviewCol && !mergeCol) {
    reviewCol = Math.max(cols + 1, 11);
    sheet.getRange(1, reviewCol).setValue('REVISION EMAIL');
    mergeCol = reviewCol + 1;
    sheet.getRange(1, mergeCol).setValue(FEST_MERGE_STATUS_HEADER);
  } else if (!reviewCol && mergeCol) {
    sheet.insertColumnBefore(mergeCol);
    reviewCol = mergeCol;
    mergeCol += 1;
    sheet.getRange(1, reviewCol).setValue('REVISION EMAIL');
    sheet.getRange(1, mergeCol).setValue(FEST_MERGE_STATUS_HEADER);
  } else if (reviewCol && !mergeCol) {
    mergeCol = Math.max(sheet.getLastColumn() + 1, reviewCol + 1);
    sheet.getRange(1, mergeCol).setValue(FEST_MERGE_STATUS_HEADER);
  }

  sheet.getRange(1, reviewCol, 1, 1)
    .setBackground('#FBC02D')
    .setFontColor('#8B0000')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange(1, mergeCol, 1, 1)
    .setValue(FEST_MERGE_STATUS_HEADER)
    .setHorizontalAlignment('center');
  return true;
}

function isFestivalSheetName_(name) {
  const n = cleanText_(name).toUpperCase();
  if (!n) return false;
  if (/^PTE[_-]/.test(n)) return true;
  return /^(URBANO?|POP|INDIE|ROCK|ELECTR(?:ONICA)?|JAZZ|FLAM(?:ENCO)?|RUMBA|MR|MC|MFR|MEC)_(S|L|XL)$/.test(n);
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
    'A partir de ahora, al abrir la hoja, aparecera el menu CRM FESTIVALES | RUBEN COTON.'
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

  let expectedPass = '';
  try {
    expectedPass = getFestSecurityPassword_();
  } catch (err) {
    ui.alert(err && err.message ? err.message : 'No se pudo validar la password de seguridad.');
    return;
  }

  if (pass !== expectedPass) {
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

function menuAuditarEmailsCRMFestivales() {
  ejecutarConPassword_(auditarEmailsCRMFestivales, 'Auditar emails + duplicados');
}

function menuAuditarEstructuraCRMFestivales() {
  ejecutarConPassword_(auditarEstructuraCRMFestivales, 'Auditar estructura');
}

function menuAuditarClasificacionCRMFestivales() {
  ejecutarConPassword_(auditarClasificacionGeneroTamanoCRMFestivales, 'Auditar genero + tamano S/L/XL');
}

function menuAuditorExtremoCRMFestivales() {
  ejecutarConPassword_(auditoriaEstresCRMFestivales, 'Modo auditor extremo (stress test)');
}

function menuInstalarTriggerCRMFestivales() {
  ejecutarConPassword_(instalarTriggerMenuCRMFestivales, 'Instalar trigger de menu');
}

function menuLimpiarTriggersCRMFestivales() {
  ejecutarConPassword_(limpiarTriggersMenuCRMFestivales, 'Limpiar triggers de menu');
}

function menuInstalarTriggerRevisionEmailsCRMFestivales() {
  ejecutarConPassword_(instalarTriggerRevisionEmailsCRMFestivales, 'Instalar trigger revision emails');
}

function menuLimpiarTriggerRevisionEmailsCRMFestivales() {
  ejecutarConPassword_(limpiarTriggerRevisionEmailsCRMFestivales, 'Limpiar trigger revision emails');
}

function menuGuiaArquitecturaCRMFestivales() {
  ejecutarConPassword_(mostrarGuiaIntegracionCRMFestivales, 'Guia de arquitectura');
}

function menuConfigurarCredencialesCRMFestivales() {
  configurarCredencialesCRMFestivales_();
}

function configurarCredencialesCRMFestivales_() {
  const ui = SpreadsheetApp.getUi();

  const keyPrompt = ui.prompt(
    'Configurar Gemini',
    'Pega la API Key de Gemini (opcional si ya existe). Se guardara en Script Properties.',
    ui.ButtonSet.OK_CANCEL
  );
  if (keyPrompt.getSelectedButton() !== ui.Button.OK) return;

  const passPrompt = ui.prompt(
    'Configurar Seguridad',
    'Define la password interna FEST_SECURITY_PASSWORD (minimo 6 caracteres).',
    ui.ButtonSet.OK_CANCEL
  );
  if (passPrompt.getSelectedButton() !== ui.Button.OK) return;

  const apiKey = cleanText_(keyPrompt.getResponseText());
  const pass = cleanText_(passPrompt.getResponseText());
  if (!pass || pass.length < 6) {
    ui.alert('La password debe tener al menos 6 caracteres.');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('FEST_SECURITY_PASSWORD', pass);
  if (apiKey) props.setProperty('FEST_GEMINI_API_KEY', apiKey);

  ui.alert('Credenciales guardadas en Script Properties.\nNo quedan expuestas en el codigo.');
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
      const normalizedRows = normalizeSheetRows_(data, sheet.getName());
      rewriteSheet_(sheet, normalizedRows);
      applyVisualDesignToSheet_(sheet);
      totalRows += normalizedRows.length;
      totalSheets += 1;
    });

    SpreadsheetApp.getUi().alert(
      'Homogeneidad aplicada.\n' +
      'Pestanas procesadas: ' + totalSheets + '\n' +
      'Filas normalizadas: ' + totalRows + '\n\n' +
      'Columnas fijadas en A:L con el mismo orden para todas.'
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
    const firstRow = sheet.getRange(1, 1, 1, Math.max(FEST_HOMO.HEADER.length, sheet.getLastColumn())).getValues()[0];
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
      if (!inOrder) msg += ' | orden distinto en A:L';
      lines.push(msg);
    }
  });

  SpreadsheetApp.getUi().alert('Auditoria de estructura\n\n' + lines.join('\n'));
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
        if (examples.length < 18) examples.push(sheet.getName() + '!A' + (r + 1) + ' -> genero=' + generoRow + ' (esperado ' + tax.genre + ')');
      }

      if (tax.size && sizeRow && sizeRow !== tax.size) {
        mismatchTamano++;
        if (examples.length < 18) examples.push(sheet.getName() + '!A' + (r + 1) + ' -> tamano=' + sizeRow + ' (esperado ' + tax.size + ')');
      }
    }
  });

  const resumen = [
    'Auditoria de clasificacion (' + FEST_ARCHITECT + ')',
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

function auditoriaEstresCRMFestivales() {
  if (!validarSesionSegura_('Modo auditor extremo (stress test)')) return;

  const started = Date.now();
  const failures = [];
  const warnings = [];
  const notes = [];
  const check = (ok, label, detail) => {
    if (!ok) failures.push(label + (detail ? ' -> ' + detail : ''));
  };

  check(sizeCodeFromAforo_('0') === 'S', 'Regla S', 'aforo 0 debe ser S');
  check(sizeCodeFromAforo_('1000') === 'S', 'Regla S', 'aforo 1000 debe ser S');
  check(sizeCodeFromAforo_('1001') === 'L', 'Regla L', 'aforo 1001 debe ser L');
  check(sizeCodeFromAforo_('9999') === 'L', 'Regla L', 'aforo 9999 debe ser L');
  check(sizeCodeFromAforo_('10000') === 'XL', 'Regla XL', 'aforo 10000 debe ser XL');
  check(formatSpanishPhone_('612345678') === '+34 612 345 678', 'Formato telefono', '612345678');
  check(isValidEmailList_('a@b.com; c@d.es') === true, 'Email list', 'lista valida');

  const fuzzInputs = [null, undefined, '', '   ', 'POP', 'MUSICA REGIONAL', '0034612345678', -1, 0, 1000, 1001, 9999, 10000, {}, [], true, false];
  const fuzzFns = [
    cleanText_, normalizeHeader_, parseAforo_, normalizeAforoForDisplay_,
    normalizeEmailCell_, normalizePhoneCell_, formatSpanishPhone_,
    normalizeContactName_, isValidEmailList_, isPlaceholderText_,
    normalizeGenreCode_, sizeCodeFromAforo_
  ];

  for (let i = 0; i < fuzzFns.length; i++) {
    for (let j = 0; j < fuzzInputs.length; j++) {
      try {
        fuzzFns[i](fuzzInputs[j]);
      } catch (err) {
        failures.push('Throw en helper: ' + (fuzzFns[i].name || 'anon') + ' con input[' + j + ']');
      }
    }
  }

  const handlers = [
    'menuHomogeneizarCRMFestivales',
    'menuAplicarDisenoCRMFestivales',
    'menuDepurarContactosCRMFestivales',
    'menuDepurarContactosGeminiCRMFestivales',
    'menuAuditarEstructuraCRMFestivales',
    'menuAuditarClasificacionCRMFestivales',
    'menuAuditorExtremoCRMFestivales',
    'menuInstalarTriggerCRMFestivales',
    'menuLimpiarTriggersCRMFestivales',
    'menuGuiaArquitecturaCRMFestivales'
  ];

  for (let i = 0; i < handlers.length; i++) {
    const fn = handlers[i];
    if (typeof this[fn] !== 'function') failures.push('Handler no encontrado: ' + fn);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getFestivalSheets_(ss);
  if (!sheets.length) warnings.push('No se detectaron hojas de festivales por patron de nombre.');

  let rowsReviewed = 0;
  let badEmails = 0;
  let badPhones = 0;
  let noAforo = 0;
  let noGenero = 0;

  for (let s = 0; s < sheets.length; s++) {
    const values = sheets[s].getDataRange().getValues();
    if (values.length < 2) continue;

    const map = buildHeaderMap_(values[0]);
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const nombre = cleanText_(valueAt_(row, map.nombre));
      if (!nombre) continue;

      rowsReviewed++;
      const email = normalizeEmailCell_(valueAt_(row, map.email));
      const phone = normalizePhoneCell_(valueAt_(row, map.telefono));
      const genre = normalizeGenreCode_(valueAt_(row, map.genero));
      const size = sizeCodeFromAforo_(valueAt_(row, map.aforo));

      if (email && !isValidEmailList_(email)) badEmails++;
      if (phone && !formatSpanishPhone_(phone)) badPhones++;
      if (!genre) noGenero++;
      if (!size) noAforo++;

      if (Date.now() - started > 250000) {
        warnings.push('Auditoria cortada por tiempo de ejecucion. Vuelve a lanzarla para terminar.');
        break;
      }
    }
  }

  notes.push('Filas revisadas: ' + rowsReviewed);
  notes.push('Emails sospechosos: ' + badEmails);
  notes.push('Telefonos sospechosos: ' + badPhones);
  notes.push('Sin genero: ' + noGenero);
  notes.push('Sin aforo/tamano: ' + noAforo);

  const status = failures.length ? 'FAIL' : 'OK';
  const message = [
    'MODO AUDITOR EXTREMO (' + status + ')',
    '',
    'Errores criticos: ' + failures.length,
    'Alertas: ' + warnings.length,
    '',
    notes.join('\n'),
    '',
    failures.length ? ('Top errores:\n- ' + failures.slice(0, 12).join('\n- ')) : 'Sin errores criticos en stress test.',
    warnings.length ? ('\n\nAlertas:\n- ' + warnings.slice(0, 8).join('\n- ')) : ''
  ].join('\n');

  SpreadsheetApp.getUi().alert(message);
}

function mostrarGuiaIntegracionCRMFestivales() {
  const html = [
    '<div style="font-family:Arial,sans-serif;padding:14px;line-height:1.5;color:#222;">',
    '<h2 style="margin-top:0;">Arquitectura CRM FESTIVALES</h2><p><b>ARQUITECTO:</b> ' + FEST_ARCHITECT + '</p>',
    '<p><b>1) En la propia hoja (Apps Script)</b><br>Extensiones -> Apps Script. Editas funciones y ejecutas desde menu o triggers.</p>',
    '<p><b>2) Con repositorio local (clasp)</b><br>Puedes sincronizar el proyecto de Apps Script con archivos .gs en tu ordenador y versionarlo con Git.</p>',
    '<p><b>3) Con APIs externas</b><br>Tu script puede llamar APIs (Gemini u otras) con UrlFetchApp y guardar resultados en celdas.</p>',
    '<p><b>4) Modificaciones seguras recomendadas</b><br>Crea copia de la hoja antes de cambios grandes, usa una pestana de pruebas, y luego aplicas a produccion.</p>',
    '<p><b>5) Flujo recomendado para ti</b><br>Menu CRM FESTIVALES -> escaner total -> depurar contactos -> auditorias de estructura y clasificacion.</p>',
    '</div>'
  ].join('');

  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(html).setWidth(520).setHeight(380),
    'Guia CRM Festivales'
  );
}

function getFestivalSheets_(ss) {
  const valid = [];
  const reMain = /^(URBANO?|POP|INDIE|ROCK|ELECTR(?:ONICA)?|JAZZ|FLAM(?:ENCO)?|RUMBA|MR|MC|MFR|MEC)_(S|L|XL)$/i;
  const rePending = /^PTE[_\-]/i;

  ss.getSheets().forEach((sheet) => {
    const name = (sheet.getName() || '').trim();
    if (reMain.test(name) || rePending.test(name)) {
      valid.push(sheet);
    }
  });

  return valid;
}

function parseSheetTaxonomy_(sheetName) {
  const name = cleanText_(sheetName).toUpperCase();
  if (/^PTE[_-]/.test(name)) return { genre: 'PTE', size: '' };

  const m = name.match(/^(URBANO?|POP|INDIE|ROCK|ELECTR(?:ONICA)?|JAZZ|FLAM(?:ENCO)?|RUMBA|MR|MC|MFR|MEC)_(S|L|XL)$/);
  if (!m) return { genre: '', size: '' };

  let genre = m[1];
  if (genre === 'URBANO') genre = 'URBAN';
  if (genre === 'ELECTRONICA') genre = 'ELECTR';
  if (genre === 'FLAMENCO') genre = 'FLAM';
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

function sizeCodeFromAforo_(aforoRaw) {
  const n = parseAforo_(aforoRaw);
  if (n === '' || isNaN(n)) return '';
  if (n <= 1000) return 'S';
  if (n >= 10000) return 'XL';
  return 'L';
}

function normalizeSheetRows_(data, sheetName) {
  if (!data || data.length === 0) return [];

  const header = data[0];
  const map = buildHeaderMap_(header);
  const out = [];
  const tax = parseSheetTaxonomy_(sheetName || '');

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const obj = {
      nombre: cleanText_(valueAt_(row, map.nombre)),
      genero: normalizeGenreCode_(valueAt_(row, map.genero)) || tax.genre || cleanText_(valueAt_(row, map.genero)),
      aforo: normalizeAforoForDisplay_(valueAt_(row, map.aforo)),
      ubicacion: normalizeMunicipioName_(valueAt_(row, map.ubicacion)),
      provincia: normalizeProvinciaName_(valueAt_(row, map.provincia)),
      ccaa: normalizeCcaaName_(valueAt_(row, map.ccaa)),
      email: normalizeEmailCell_(valueAt_(row, map.email)),
      telefono: normalizePhoneCell_(valueAt_(row, map.telefono)),
      contacto: normalizeContactName_(valueAt_(row, map.contacto)),
      observaciones: cleanText_(valueAt_(row, map.observaciones)),
      revisionEmail: normalizeRevisionStatus_(valueAt_(row, map.reviewEmail)),
      mergeStatus: cleanText_(valueAt_(row, map.mergeStatus))
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
      obj.observaciones,
      obj.revisionEmail,
      obj.mergeStatus
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

  const tax = parseSheetTaxonomy_(sheet.getName());
  const tabColor = tax.genre === 'PTE' ? '#FBC02D' : '#C00000';
  try { sheet.setTabColor(tabColor); } catch (err) {}

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
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 170);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 260);
  sheet.setColumnWidth(8, 140);
  sheet.setColumnWidth(9, 190);
  sheet.setColumnWidth(10, 260);
  sheet.setColumnWidth(11, 190);
  sheet.setColumnWidth(12, 150);

  if (lastRow > 1) {
    const map = buildHeaderMap_(sheet.getRange(1, 1, 1, lastCol).getValues()[0]);
    const reviewCol = map.reviewEmail + 1;
    const mergeCol = map.mergeStatus + 1;
    const reviewColLetter = columnNumberToLetter_(reviewCol);

    sheet.getRange(2, 3, lastRow - 1, 1).setHorizontalAlignment('center');
    sheet.getRange(2, 2, lastRow - 1, 1).setHorizontalAlignment('center');
    sheet.getRange(2, 7, lastRow - 1, 2).setHorizontalAlignment('left');

    const generoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(FEST_GENRE_DROPDOWN, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 2, lastRow - 1, 1).setDataValidation(generoRule);
    sheet.getRange(2, reviewCol, lastRow - 1, 1)
      .setDataValidation(buildEmailReviewValidationRule_())
      .setHorizontalAlignment('center')
      .setFontWeight('bold');
    sheet.getRange(1, reviewCol, 1, 1)
      .setBackground('#FBC02D')
      .setFontColor('#8B0000')
      .setHorizontalAlignment('center')
      .setFontWeight('bold');
    sheet.getRange(1, mergeCol, 1, 1).setValue(FEST_MERGE_STATUS_HEADER).setHorizontalAlignment('center');
    sheet.getRange(2, mergeCol, lastRow - 1, 1).clearDataValidations().setHorizontalAlignment('left').setFontWeight('normal');

    const rangeForRules = sheet.getRange(2, 1, Math.max(1, lastRow - 1), lastCol);

    const bienRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"";$' + reviewColLetter + '2="BIEN")')
      .setBackground('#D9EAD3')
      .setRanges([rangeForRules])
      .build();

    const corregidoRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"";$' + reviewColLetter + '2="CORREGIDO")')
      .setBackground('#D9E8FB')
      .setRanges([rangeForRules])
      .build();

    const malRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"";$' + reviewColLetter + '2="MAL")')
      .setBackground('#FDE2E2')
      .setRanges([rangeForRules])
      .build();

    sheet.setConditionalFormatRules([bienRule, corregidoRule, malRule]);
  }
}

function buildHeaderMap_(headerRow) {
  const safeHeader = Array.isArray(headerRow) ? headerRow : [];
  const map = {};
  for (let c = 0; c < safeHeader.length; c++) {
    const h = normalizeHeader_(safeHeader[c]);
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
    if ((h === 'REVISION EMAIL' || h === 'ESTADO EMAIL' || h === 'EMAIL REVISADO') && map.reviewEmail === undefined) map.reviewEmail = c;
    if ((h === 'MERGE STATUS' || h === 'MERGESTATUS' || h === 'MERGE') && map.mergeStatus === undefined) map.mergeStatus = c;
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
  if (map.reviewEmail === undefined && map.mergeStatus !== undefined) map.reviewEmail = map.mergeStatus + 1;
  if (map.reviewEmail !== undefined && map.mergeStatus === undefined) map.mergeStatus = map.reviewEmail + 1;
  if (map.reviewEmail === undefined) map.reviewEmail = 10;
  if (map.mergeStatus === undefined) map.mergeStatus = 11;

  if (map.reviewEmail === map.mergeStatus) {
    const idx = Math.max(0, Number(map.reviewEmail) || 0);
    const h = normalizeHeader_(valueAt_(safeHeader, idx));
    if (h === 'MERGE STATUS' || h === 'MERGESTATUS' || h === 'MERGE') {
      map.mergeStatus = idx;
      map.reviewEmail = idx + 1;
    } else {
      map.reviewEmail = idx;
      map.mergeStatus = idx + 1;
    }
  }

  map.reviewEmail = Math.max(0, Number(map.reviewEmail) || 0);
  map.mergeStatus = Math.max(0, Number(map.mergeStatus) || 0);

  return map;
}

function buildEmailReviewValidationRule_() {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(FEST_EMAIL_REVIEW_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
}

function normalizeRevisionStatus_(v) {
  const t = cleanText_(v).toUpperCase();
  if (!t) return 'MAL';
  for (let i = 0; i < FEST_EMAIL_REVIEW_OPTIONS.length; i++) {
    if (t === FEST_EMAIL_REVIEW_OPTIONS[i]) return t;
  }
  if (t === 'OK_VERIFICADO_WEB' || t === 'VERIFICADO' || t === 'CORRECTO') return 'BIEN';
  if (t === 'CAMBIO' || t === 'ACTUALIZADO' || t === 'CAMBIADO') return 'CORREGIDO';
  if (t === 'PENDIENTE_REVISION' || t === 'REVISAR_WEB' || t === 'CORREGIR_EMAIL' || t === 'SIN_EMAIL' || t === 'DUPLICADO') return 'MAL';
  return 'MAL';
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

function columnNumberToLetter_(col) {
  let n = Math.max(1, Number(col) || 1);
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
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
    obj.observaciones ||
    obj.revisionEmail ||
    obj.mergeStatus
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

function normalizeMunicipioName_(v) {
  return normalizeGeoTitle_(v);
}

function normalizeProvinciaName_(v) {
  const key = normalizeGeoKey_(v);
  const map = {
    'A CORUNA': 'A Coruña',
    'ALAVA': 'Álava',
    'AVILA': 'Ávila',
    'CADIZ': 'Cádiz',
    'CASTELLON': 'Castellón',
    'CORDOBA': 'Córdoba',
    'GIPUZKOA': 'Gipuzkoa',
    'GUIPUZCOA': 'Gipuzkoa',
    'JAEN': 'Jaén',
    'LA CORUNA': 'A Coruña',
    'LEON': 'León',
    'MALAGA': 'Málaga',
    'ORENSE': 'Ourense',
    'VIZCAYA': 'Bizkaia'
  };
  return map[key] || normalizeGeoTitle_(v);
}

function normalizeCcaaName_(v) {
  const key = normalizeGeoKey_(v);
  const map = {
    'ANDALUCIA': 'Andalucía',
    'ARAGON': 'Aragón',
    'ASTURIAS': 'Asturias',
    'CANARIAS': 'Canarias',
    'CASTILLA LA MANCHA': 'Castilla-La Mancha',
    'CASTILLA Y LEON': 'Castilla y León',
    'CATALUNA': 'Cataluña',
    'COMUNIDAD VALENCIANA': 'Comunidad Valenciana',
    'COMUNITAT VALENCIANA': 'Comunidad Valenciana',
    'C VALENCIANA': 'Comunidad Valenciana',
    'LA RIOJA': 'La Rioja',
    'MADRID': 'Comunidad de Madrid',
    'COMUNIDAD DE MADRID': 'Comunidad de Madrid',
    'MURCIA': 'Región de Murcia',
    'REGION DE MURCIA': 'Región de Murcia',
    'PAIS VASCO': 'País Vasco',
    'EUSKADI': 'País Vasco'
  };
  return map[key] || normalizeGeoTitle_(v);
}

function normalizeGeoTitle_(v) {
  const txt = cleanText_(v).toLowerCase();
  if (!txt) return '';
  const keepLower = { de: true, del: true, la: true, las: true, los: true, y: true, e: true };
  const parts = txt.split(' ');
  const out = [];
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    if (!p) continue;
    if (i > 0 && keepLower[p]) out.push(p);
    else out.push(p.charAt(0).toUpperCase() + p.slice(1));
  }
  return out.join(' ');
}

function normalizeGeoKey_(v) {
  return cleanText_(v)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
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

function buildGeminiResponseCacheKey_(prompt, responseSchema) {
  try {
    const base = String(prompt || '') + '|' + JSON.stringify(responseSchema || {});
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, base);
    const hex = digest.map((b) => ((b < 0 ? b + 256 : b).toString(16).padStart(2, '0'))).join('');
    return 'FEST_GEM_CACHE_' + hex.substring(0, 64);
  } catch (err) {
    return '';
  }
}

function invocarGeminiConFallback_(prompt, responseSchema) {
  const apiKey = getFestGeminiApiKey_();
  if (!apiKey) return null;

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
      maxOutputTokens: FEST_GEMINI_BASE_MAX_OUTPUT_TOKENS,
      responseMimeType: 'application/json',
      responseSchema: responseSchema
    }
  };

  const cacheKey = buildGeminiResponseCacheKey_(prompt, responseSchema);
  if (cacheKey) {
    try {
      const hit = CacheService.getScriptCache().get(cacheKey);
      if (hit) return { model: 'cache', data: JSON.parse(hit) };
    } catch (errCacheRead) {}
  }

  const baseOptions = {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'CRM-FESTIVALES/1.0' }
  };

  for (let i = 0; i < FEST_GEMINI_MODELS.length; i++) {
    const model = FEST_GEMINI_MODELS[i];
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + apiKey;

    for (let attempt = 0; attempt < FEST_GEMINI_MAX_RETRIES_PER_MODEL; attempt++) {
      try {
        const maxTokens = (attempt > 0)
          ? Math.max(FEST_GEMINI_BASE_MAX_OUTPUT_TOKENS, FEST_GEMINI_RETRY_MAX_OUTPUT_TOKENS)
          : FEST_GEMINI_BASE_MAX_OUTPUT_TOKENS;
        payload.generationConfig.maxOutputTokens = maxTokens;

        const options = Object.assign({}, baseOptions, {
          payload: JSON.stringify(payload)
        });

        const res = UrlFetchApp.fetch(url, options);
        const code = res.getResponseCode();
        const raw = res.getContentText();

        if (code === 200) {
          const parsed = parseGeminiJson_(raw);
          if (parsed) {
            if (cacheKey) {
              try {
                CacheService.getScriptCache().put(cacheKey, JSON.stringify(parsed), FEST_GEMINI_RESPONSE_CACHE_TTL_SEC);
              } catch (errCacheWrite) {}
            }
            return { model: model, data: parsed };
          }

          if (attempt < FEST_GEMINI_MAX_RETRIES_PER_MODEL - 1) {
            Utilities.sleep(650 * Math.pow(2, attempt));
            continue;
          }

          break;
        }

        if (code === 401 || code === 403) return null;
        if (code === 404) break;
        if (code === 429 || code >= 500) {
          Utilities.sleep(650 * Math.pow(2, attempt));
          continue;
        }

        break;
      } catch (err) {
        Utilities.sleep(650 * Math.pow(2, attempt));
      }
    }
  }

  return null;
}

function extractGeminiTextFromCandidate_(candidate) {
  const parts = ((candidate && candidate.content) ? candidate.content.parts : []) || [];
  const texts = [];
  for (let i = 0; i < parts.length; i++) {
    const t = cleanText_(parts[i] && parts[i].text);
    if (t) texts.push(t);
  }
  return texts.join('\n').trim();
}

function parseGeminiJson_(rawText) {
  try {
    const root = JSON.parse(rawText);
    const candidates = Array.isArray(root && root.candidates) ? root.candidates : [];
    if (!candidates.length) return null;

    for (let c = 0; c < candidates.length; c++) {
      const cleaned = String(extractGeminiTextFromCandidate_(candidates[c]) || '')
        .replace(/`{3}json/gi, '')
        .replace(/`{3}/g, '')
        .trim();

      if (!cleaned) continue;

      try {
        if (cleaned.charAt(0) === '{' || cleaned.charAt(0) === '[') {
          const parsedDirect = JSON.parse(cleaned);
          if (Array.isArray(parsedDirect)) return parsedDirect.length ? parsedDirect[0] : null;
          return parsedDirect;
        }

        const match = cleaned.match(/(\{[\s\S]*\}|\[[\s\S]*\])/);
        if (!match) continue;

        const parsed = JSON.parse(match[0]);
        if (Array.isArray(parsed)) return parsed.length ? parsed[0] : null;
        return parsed;
      } catch (innerErr) {
        // sigue con siguiente candidato
      }
    }

    return null;
  } catch (err) {
    return null;
  }
}
