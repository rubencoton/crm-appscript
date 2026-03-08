// =============================================================================
// CRM FESTIVALES - MENU OPERATIVO + HOMOGENEIDAD + DISENO
// Listo para Google Apps Script
// =============================================================================

// =============================================================================
// MAPA FUNCIONAL DEL SCRIPT (CRM FESTIVALES)
// =============================================================================
// Este proyecto trabaja sobre Google Sheets + Apps Script para CRM de festivales.
//
// Estructura canonica de columnas (A:J):
// A NOMBRE FESTIVAL
// B GENERO
// C AFORO
// D UBICACION (MUNICIPIO)
// E PROVINCIA
// F CCAA
// G EMAIL
// H TELEFONO
// I NOMBRE CONTACTO
// J OBSERVACIONES
//
// Convencion de pestanas:
// - <GENERO>_<TAMANO> -> ej. URBAN_S, POP_L, ELECTR_XL
// - PTE_* para pendientes de clasificar
// - Alias historicos MR/MC se normalizan a MFR/MEC
//
// Regla de tamano por aforo:
// - S: 0..1000 (incluido 1000)
// - L: 1001..9999
// - XL: 10000..infinito
//
// Modulos principales:
// 1) Menu y seguridad por password
// 2) Homogeneizacion de estructura + diseno
// 3) Depuracion local de contactos
// 4) Depuracion con Gemini (fallback + cache)
// 5) Auditorias de estructura, clasificacion y stress test
// =============================================================================

// =============================================================================
// CONTEXTO DEL CHAT Y ACUERDOS OPERATIVOS (MEMORIA PARA NUEVOS CHATS)
// =============================================================================
// Este bloque deja por escrito las decisiones tomadas durante el desarrollo:
//
// 1) Alcance actual:
// - CRM de festivales para Espana (fase actual del proyecto).
// - Prioridad de calidad de contacto: EMAIL, TELEFONO y NOMBRE CONTACTO.
//
// 2) Clasificacion de genero acordada:
// - URBAN, POP, INDIE, ROCK, ELECTR, JAZZ, FLAM, RUMBA, MEC (musica clasica),
//   MFR (musica regional), y PTE para pendientes/no clasificados.
//
// 3) Regla de tamano por aforo acordada:
// - S: desde 0 hasta 1000 (1000 incluido)
// - L: desde 1001 hasta 9999
// - XL: desde 10000 hasta infinito
//
// 4) Experiencia visual acordada:
// - Menu principal con emojis.
// - Pestanas coloreadas por genero.
// - Tabla homogenea en orden A:J y reglas visuales de alerta.
//
// 5) Seguridad acordada:
// - Las acciones de menu se ejecutan bajo password.
// - Password actual configurada en este script: +rubencoton26
//
// 6) IA/Gemini acordado:
// - Modelo preferente: gemini-3.1-pro-preview (maxima prioridad).
// - Fallback automatico a otros modelos en cascada.
// - Cache de respuestas para reducir coste y latencia.
//
// 7) Arquitectura y sincronizacion acordadas:
// - Desarrollo y versionado en GitHub (rama main).
// - Despliegue en Google Apps Script via clasp (gas:push).
// - Este archivo .gs es la fuente principal operativa del CRM FESTIVALES.
//
// 8) Orden recomendado de ejecucion en la hoja:
// - 1) Escaner total + homogeneizar
// - 2) Depurar contactos local
// - 3) Depurar contactos con Gemini (por lotes)
// - 4) Auditar estructura
// - 5) Auditar genero + tamano
// - 6) Modo auditor extremo cuando se quiera stress test
//
// 9) Nota de continuidad:
// - Si abres otro chat nuevo, pide revisar primero este bloque de memoria y
//   despues ejecutar la auditoria antes de tocar reglas de negocio.
// =============================================================================
/**
 * Configuracion central del CRM de festivales.
 * - HEADER define el contrato de columnas fijo para todas las pestanas.
 * - COLORS define la identidad visual y codigos de alerta.
 */
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

// Seguridad y IA:
// - FEST_SECURITY_PASSWORD protege acciones de menu sensibles.
// - FEST_GEMINI_API_KEY y FEST_GEMINI_MODELS habilitan depuracion asistida por IA.
// - El orden de modelos define fallback: se intenta del mas potente al mas rapido.
const FEST_SECURITY_PASSWORD = '+rubencoton26';
const FEST_GEMINI_API_KEY = 'AIzaSyC2AnnQuFgKOR_qGNl4jTrsoWF672bnK0M';
const FEST_GEMINI_MODELS = [
  'gemini-3.1-pro-preview',
  'gemini-2.5-pro',
  'gemini-3-flash-preview',
  'gemini-2.5-flash',
  'gemini-2.5-flash-lite',
  'gemini-flash-latest'
];
const FEST_GEMINI_MAX_RETRIES_PER_MODEL = 2;
const FEST_GEMINI_RESPONSE_CACHE_TTL_SEC = 6 * 60 * 60;
const FEST_MAX_RUNTIME_MS = 4.7 * 60 * 1000;
const FEST_ARCHITECT = 'RUBEN COTON';
const FEST_GENRE_DROPDOWN = [
  '🧢 URBAN', '🎤 POP', '🎸 INDIE', '🤘 ROCK', '🎛️ ELECTR',
  '🎷 JAZZ', '💃 FLAM', '🪘 RUMBA', '🎼 MEC', '🌄 MFR'
];

// Si no tienes otro onOpen en el proyecto, este ya deja el menu automatico.

/**
 * Hook nativo de Google Sheets.
 * Se ejecuta al abrir el archivo y dibuja el menu operativo del CRM.
 */
function onOpen(e) {
  crearMenuCRMFestivales_();
}

// Alias para compatibilidad con versiones previas del script.
function onOpenFestivalesHomogeneidad_() {
  crearMenuCRMFestivales_();
}

/**
 * Construye el menu principal visible en la barra de Google Sheets.
 * Todas las opciones llaman wrappers protegidos por password.
 */
function crearMenuCRMFestivales_() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 CRM FESTIVALES | RUBEN COTON')
    .addItem('🚀 Escaner total + homogeneizar (seguro)', 'menuHomogeneizarCRMFestivales')
    .addItem('🎨 Solo armonizar diseno visual (seguro)', 'menuAplicarDisenoCRMFestivales')
    .addItem('📧 Depurar contactos local (seguro)', 'menuDepurarContactosCRMFestivales')
    .addItem('🧠 Depurar contactos con Gemini (seguro)', 'menuDepurarContactosGeminiCRMFestivales')
    .addItem('🛰️ Auditar estructura (seguro)', 'menuAuditarEstructuraCRMFestivales')
    .addItem('🧭 Auditar genero + tamano S/L/XL (seguro)', 'menuAuditarClasificacionCRMFestivales')
    .addItem('💥 Modo auditor extremo (stress test)', 'menuAuditorExtremoCRMFestivales')
    .addSeparator()
    .addItem('⚙️ Instalar trigger de menu (seguro)', 'menuInstalarTriggerCRMFestivales')
    .addItem('🧹 Limpiar triggers de menu (seguro)', 'menuLimpiarTriggersCRMFestivales')
    .addSeparator()
    .addItem('📚 Guia de arquitectura (seguro)', 'menuGuiaArquitecturaCRMFestivales')
    .addToUi();
}

/**
 * Instala un trigger onOpen dedicado para asegurar que el menu aparece siempre,
 * incluso tras copiar el archivo o moverlo de carpeta/proyecto.
 */
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
    'A partir de ahora, al abrir la hoja, aparecera el menu 🚀 CRM FESTIVALES | RUBEN COTON.'
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


/**
 * Wrapper de seguridad para menu.
 * 1) Solicita password
 * 2) Activa sesion corta en cache
 * 3) Ejecuta la accion pedida
 */
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

/**
 * Marca una sesion temporal de confianza (TTL corto) en cache de usuario.
 */
function activarSesionSegura_() {
  try {
    CacheService.getUserCache().put('FEST_AUTH_OK', '1', 300);
  } catch (err) {
    // no-op
  }
}

/**
 * Verifica que la accion se lanzo desde el wrapper con password.
 * Evita ejecucion directa accidental desde editor sin autenticacion previa.
 */
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

// Wrappers de menu: unifican el paso de seguridad para cada modulo.
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

function menuGuiaArquitecturaCRMFestivales() {
  ejecutarConPassword_(mostrarGuiaIntegracionCRMFestivales, 'Guia de arquitectura');
}

/**
 * 1) Unifica TODAS las pestanas de festivales al mismo orden de columnas.
 * 2) Aplica diseno visual consistente para una lectura comoda.
 */
/**
 * MODULO PRINCIPAL DE NORMALIZACION.
 * - Recorre todas las pestanas de festivales validas.
 * - Reescribe cada hoja a formato canonico A:J.
 * - Normaliza genero/aforo/contactos y reaplica diseno visual.
 * - Usa LockService para evitar doble ejecucion simultanea.
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
      'Columnas fijadas en A:J con el mismo orden para todas.'
    );
  } finally {
    lock.releaseLock();
  }
}

/**
 * Reaplica solo el formato visual a las pestanas detectadas.
 */
/**
 * Solo reaplica formato visual corporativo sin tocar datos.
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
/**
 * Depuracion local (sin IA):
 * - EMAIL: limpia, deduplica y valida sintaxis
 * - TELEFONO: intenta normalizar a +34 XXX XXX XXX
 * - CONTACTO: capitalizacion y limpieza basica
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

/**
 * Auditoria de estructura:
 * comprueba presencia de columnas requeridas y orden canonico en A:J.
 */
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

  SpreadsheetApp.getUi().alert('🛰️ Auditoria de estructura\n\n' + lines.join('\n'));
}

/**
 * Auditoria de coherencia de clasificacion:
 * - Compara genero de fila vs genero esperado por nombre de pestana.
 * - Compara tamano calculado por aforo vs sufijo S/L/XL de la pestana.
 */
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

/**
 * Stress test interno del motor:
 * - Pruebas de borde de reglas S/L/XL
 * - Fuzz sobre helpers de normalizacion
 * - Chequeo de handlers de menu y salud basica de datos
 */
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

  const fuzzInputs = [null, undefined, '', '   ', '🎤 POP', 'MUSICA REGIONAL', '0034612345678', -1, 0, 1000, 1001, 9999, 10000, {}, [], true, false];
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
    '💥 MODO AUDITOR EXTREMO (' + status + ')',
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

/**
 * Dialogo de ayuda para explicar arquitectura local + Apps Script + GitHub.
 */
function mostrarGuiaIntegracionCRMFestivales() {
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

/**
 * Devuelve solo pestanas de trabajo del CRM FESTIVALES.
 * Ignora hojas auxiliares que no cumplan la taxonomia de nombre.
 */
function getFestivalSheets_(ss) {
  const valid = [];
  const reMain = /^(URBAN|POP|INDIE|ROCK|ELECTR|JAZZ|FLAM|RUMBA|MR|MC|MFR|MEC)_(S|L|XL)$/i;
  const rePending = /^PTE[_\-]/i;

  ss.getSheets().forEach((sheet) => {
    const name = (sheet.getName() || '').trim();
    if (reMain.test(name) || rePending.test(name)) {
      valid.push(sheet);
    }
  });

  return valid;
}

/**
 * Extrae genero y tamano desde el nombre de pestana.
 * Ejemplo: "ELECTR_XL" -> { genre: "ELECTR", size: "XL" }
 */
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

/**
 * Canoniza genero a codigo corto interno:
 * URBAN, POP, INDIE, ROCK, ELECTR, JAZZ, FLAM, RUMBA, MEC, MFR.
 */
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

/**
 * Aplica reglas de negocio del tamano por aforo.
 * S <= 1000, L 1001..9999, XL >= 10000.
 */
function sizeCodeFromAforo_(aforoRaw) {
  const n = parseAforo_(aforoRaw);
  if (n === '' || isNaN(n)) return '';
  if (n <= 1000) return 'S';
  if (n >= 10000) return 'XL';
  return 'L';
}

/**
 * Normaliza filas de una hoja a la estructura canonica del CRM.
 * Descarta filas completamente vacias y conserva solo informacion util.
 */
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

/**
 * Reescribe la hoja completa con cabecera estandar y filas normalizadas.
 */
function rewriteSheet_(sheet, rows) {
  sheet.clear();
  sheet.getRange(1, 1, 1, FEST_HOMO.HEADER.length).setValues([FEST_HOMO.HEADER]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, FEST_HOMO.HEADER.length).setValues(rows);
  }
}

/**
 * Aplica experiencia visual uniforme:
 * - Cabecera corporativa
 * - Colores por pestana/genero
 * - Filtro, anchos, alineaciones, validaciones y formato condicional
 */
function applyVisualDesignToSheet_(sheet) {
  const lastRow = Math.max(2, sheet.getLastRow());
  const lastCol = FEST_HOMO.HEADER.length;

  sheet.setConditionalFormatRules([]);
  sheet.getBandings().forEach((b) => b.remove());

  const tax = parseSheetTaxonomy_(sheet.getName());
  const tabColors = { URBAN: '#FB8C00', POP: '#EC407A', INDIE: '#546E7A', ROCK: '#E53935', ELECTR: '#00ACC1', JAZZ: '#3949AB', FLAM: '#D81B60', RUMBA: '#43A047', MEC: '#6D4C41', MFR: '#8D6E63', PTE: '#757575' };
  if (tabColors[tax.genre]) {
    try { sheet.setTabColor(tabColors[tax.genre]); } catch (err) {}
  }

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

/**
 * Crea mapa de indices de columna tolerante a variantes de cabecera.
 * Si falta una cabecera, aplica fallback por posicion A:J.
 */
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

/**
 * Normaliza cabeceras para comparaciones robustas:
 * trim + uppercase + sin acentos + espacios colapsados.
 */
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

/**
 * Limpia y estandariza emails multiples.
 * Devuelve lista separada por "; " con deduplicacion.
 */
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

/**
 * Intenta extraer el primer telefono espanol valido de una celda.
 */
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

/**
 * Formatea telefono espanol a +34 XXX XXX XXX.
 * Acepta variantes con prefijo 0034 o 34.
 */
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

/**
 * Capitaliza nombre de contacto manteniendo conectores comunes en minuscula.
 */
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




/**
 * Depuracion avanzada con Gemini por lotes.
 * Solo procesa filas donde la validacion local detecta debilidad de datos.
 * Respeta limite de tiempo de Apps Script para evitar cortes bruscos.
 */
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

/**
 * Decide si una fila necesita IA o si la calidad local ya es suficiente.
 */
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

/**
 * Construye prompt + schema estricto para depurar una fila concreta.
 */
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

/**
 * Firma hash de peticion para cachear respuestas de IA y ahorrar coste/tiempo.
 */
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

/**
 * Cliente Gemini con fallback por modelos:
 * - Reintentos por modelo para 429/5xx
 * - Salto de modelo ante 404
 * - Cache de respuestas para prompts repetidos
 */
function invocarGeminiConFallback_(prompt, responseSchema) {
  const apiKey = cleanText_(FEST_GEMINI_API_KEY);
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
      maxOutputTokens: 800,
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

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'CRM-FESTIVALES/1.0' }
  };

  for (let i = 0; i < FEST_GEMINI_MODELS.length; i++) {
    const model = FEST_GEMINI_MODELS[i];
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + apiKey;

    for (let attempt = 0; attempt < FEST_GEMINI_MAX_RETRIES_PER_MODEL; attempt++) {
      try {
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

/**
 * Parser tolerante de respuesta de Gemini.
 * Soporta JSON puro, bloques markdown y texto envolvente con JSON embebido.
 */
function parseGeminiJson_(rawText) {
  try {
    const root = JSON.parse(rawText);
    if (!root || !root.candidates || !root.candidates.length) return null;
    const part = root.candidates[0].content && root.candidates[0].content.parts
      ? root.candidates[0].content.parts[0]
      : null;
    if (!part || !part.text) return null;

    const cleaned = String(part.text)
      .replace(/`{3}json/gi, '')
      .replace(/`{3}/g, '')
      .trim();

    if (!cleaned) return null;

    if (cleaned.charAt(0) === '{' || cleaned.charAt(0) === '[') {
      const parsedDirect = JSON.parse(cleaned);
      if (Array.isArray(parsedDirect)) return parsedDirect.length ? parsedDirect[0] : null;
      return parsedDirect;
    }

    const match = cleaned.match(/(\{[\s\S]*\}|\[[\s\S]*\])/);
    if (!match) return null;

    const parsed = JSON.parse(match[0]);
    if (Array.isArray(parsed)) return parsed.length ? parsed[0] : null;
    return parsed;
  } catch (err) {
    return null;
  }
}

