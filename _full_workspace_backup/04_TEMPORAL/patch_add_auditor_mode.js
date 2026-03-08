const fs = require('fs');
const file = 'C:/Users/elrub/Desktop/CARPETA CODEX/01_PROYECTOS/festivales-github/Código.js';
let text = fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n');

function mustReplace(oldText, newText, label) {
  if (!text.includes(oldText)) throw new Error('No se encontro bloque: ' + label);
  text = text.replace(oldText, newText);
}

mustReplace(
"    .addItem('🧭 Auditar genero + tamano S/L/XL (seguro)', 'menuAuditarClasificacionCRMFestivales')\n",
"    .addItem('🧭 Auditar genero + tamano S/L/XL (seguro)', 'menuAuditarClasificacionCRMFestivales')\n    .addItem('💥 Modo auditor extremo (stress test)', 'menuAuditorExtremoCRMFestivales')\n",
'menu item auditor extremo'
);

mustReplace(
`function menuAuditarClasificacionCRMFestivales() {
  ejecutarConPassword_(auditarClasificacionGeneroTamanoCRMFestivales, 'Auditar genero + tamano S/L/XL');
}
`,
`function menuAuditarClasificacionCRMFestivales() {
  ejecutarConPassword_(auditarClasificacionGeneroTamanoCRMFestivales, 'Auditar genero + tamano S/L/XL');
}

function menuAuditorExtremoCRMFestivales() {
  ejecutarConPassword_(auditoriaEstresCRMFestivales, 'Modo auditor extremo (stress test)');
}
`,
'wrapper auditor extremo'
);

const anchor = `function mostrarGuiaIntegracionCRMFestivales() {`;
if (!text.includes('function auditoriaEstresCRMFestivales() {')) {
  const block = `function auditoriaEstresCRMFestivales() {
  if (!validarSesionSegura_('Modo auditor extremo (stress test)')) return;

  const started = Date.now();
  const failures = [];
  const warnings = [];
  const notes = [];

  const check = (ok, label, detail) => {
    if (!ok) failures.push(label + (detail ? ' -> ' + detail : ''));
  };

  // Tests deterministas de limites
  check(sizeCodeFromAforo_('0') === 'S', 'Regla S', 'aforo 0 debe ser S');
  check(sizeCodeFromAforo_('1000') === 'S', 'Regla S', 'aforo 1000 debe ser S');
  check(sizeCodeFromAforo_('1001') === 'L', 'Regla L', 'aforo 1001 debe ser L');
  check(sizeCodeFromAforo_('9999') === 'L', 'Regla L', 'aforo 9999 debe ser L');
  check(sizeCodeFromAforo_('10000') === 'XL', 'Regla XL', 'aforo 10000 debe ser XL');
  check(formatSpanishPhone_('612345678') === '+34 612 345 678', 'Formato telefono', '612345678');
  check(isValidEmailList_('a@b.com; c@d.es') === true, 'Email list', 'lista valida');

  // Fuzz corto para detectar throws en helpers puros
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

  // Sanidad de handlers de menu
  const handlers = [
    'menuHomogeneizarCRMFestivales',
    'menuAplicarDisenoCRMFestivales',
    'menuDepurarContactosCRMFestivales',
    'menuDepurarContactosGeminiCRMFestivales',
    'menuAuditarEstructuraCRMFestivales',
    'menuAuditarClasificacionCRMFestivales',
    'menuAuditorExtremoCRMFestivales',
    'instalarTriggerMenuCRMFestivales',
    'limpiarTriggersMenuCRMFestivales',
    'mostrarGuiaIntegracionCRMFestivales'
  ];
  for (let i = 0; i < handlers.length; i++) {
    const fn = handlers[i];
    if (typeof this[fn] !== 'function') {
      failures.push('Handler no encontrado: ' + fn);
    }
  }

  // Auditoria de datos real sobre hojas
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getFestivalSheets_(ss);
  if (!sheets.length) {
    warnings.push('No se detectaron hojas de festivales por patron de nombre.');
  }

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

`;
  text = text.replace(anchor, block + anchor);
}

fs.writeFileSync(file, text, 'utf8');
console.log('OK');
