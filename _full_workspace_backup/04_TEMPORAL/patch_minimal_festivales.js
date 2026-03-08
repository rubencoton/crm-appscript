const fs = require('fs');
const path = 'C:/Users/elrub/Desktop/CARPETA CODEX/01_PROYECTOS/festivales-github/Código.js';
let text = fs.readFileSync(path, 'utf8').replace(/\r\n/g, '\n');

function mustReplace(oldText, newText, label) {
  if (!text.includes(oldText)) throw new Error('No se encontro bloque: ' + label);
  text = text.replace(oldText, newText);
}

mustReplace(
  "const FEST_MAX_RUNTIME_MS = 4.7 * 60 * 1000;\n",
  "const FEST_MAX_RUNTIME_MS = 4.7 * 60 * 1000;\nconst FEST_ARCHITECT = 'RUBEN COTON';\nconst FEST_GENRE_DROPDOWN = [\n  '🧢 URBAN', '🎤 POP', '🎸 INDIE', '🤘 ROCK', '🎛️ ELECTR',\n  '🎷 JAZZ', '💃 FLAM', '🪘 RUMBA', '🎼 MEC', '🌄 MFR'\n];\n",
  'constantes'
);

mustReplace(
`function crearMenuCRMFestivales_() {
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
`,
`function crearMenuCRMFestivales_() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 CRM FESTIVALES | RUBEN COTON')
    .addItem('🚀 Escaner total + homogeneizar (seguro)', 'menuHomogeneizarCRMFestivales')
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
`,
  'menu'
);

mustReplace(
`function menuAuditarEstructuraCRMFestivales() {
  ejecutarConPassword_(auditarEstructuraCRMFestivales, 'Auditar estructura');
}
`,
`function menuAuditarEstructuraCRMFestivales() {
  ejecutarConPassword_(auditarEstructuraCRMFestivales, 'Auditar estructura');
}

function menuAuditarClasificacionCRMFestivales() {
  ejecutarConPassword_(auditarClasificacionGeneroTamanoCRMFestivales, 'Auditar genero + tamano S/L/XL');
}
`,
  'wrapper auditoria clasificacion'
);

mustReplace(
"A partir de ahora, al abrir la hoja, aparecera el menu CRM Festivales.",
"A partir de ahora, al abrir la hoja, aparecera el menu 🚀 CRM FESTIVALES | RUBEN COTON.",
'mensaje trigger'
);

mustReplace(
"SpreadsheetApp.getUi().alert('Auditoria de estructura\\n\\n' + lines.join('\\n'));",
"SpreadsheetApp.getUi().alert('🛰️ Auditoria de estructura\\n\\n' + lines.join('\\n'));",
'auditoria estructura alert'
);

mustReplace(
`function mostrarGuiaIntegracionCRMFestivales() {
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
`,
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
`,
'guia'
);

mustReplace(
"const reMain = /^(URBAN|POP|INDIE|ROCK|ELECTR|JAZZ|FLAM|RUMBA|MR|MC)_(S|L|XL)$/i;",
"const reMain = /^(URBAN|POP|INDIE|ROCK|ELECTR|JAZZ|FLAM|RUMBA|MR|MC|MFR|MEC)_(S|L|XL)$/i;",
'regex hojas'
);

mustReplace(
"function normalizeSheetRows_(data) {",
`function parseSheetTaxonomy_(sheetName) {
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

function sizeCodeFromAforo_(aforoRaw) {
  const n = parseAforo_(aforoRaw);
  if (n === '' || isNaN(n)) return '';
  if (n <= 1000) return 'S';
  if (n >= 10000) return 'XL';
  return 'L';
}

function normalizeSheetRows_(data) {`,
'helpers taxonomia'
);

mustReplace(
"      genero: cleanText_(valueAt_(row, map.genero)),",
"      genero: normalizeGenreCode_(valueAt_(row, map.genero)) || cleanText_(valueAt_(row, map.genero)) || parseSheetTaxonomy_((data[0] && data[0][0]) || '').genre,",
'normalizar genero'
);

mustReplace(
"  sheet.setFrozenRows(1);",
"  sheet.setFrozenRows(1);\n  sheet.setFrozenColumns(1);",
'freeze columns'
);

mustReplace(
`  if (lastRow > 1) {
    sheet.getRange(2, 3, lastRow - 1, 1).setHorizontalAlignment('center');
    sheet.getRange(2, 2, lastRow - 1, 1).setHorizontalAlignment('center');
    sheet.getRange(2, 7, lastRow - 1, 2).setHorizontalAlignment('left');
  }
`,
`  if (lastRow > 1) {
    sheet.getRange(2, 3, lastRow - 1, 1).setHorizontalAlignment('center');
    sheet.getRange(2, 2, lastRow - 1, 1).setHorizontalAlignment('center');
    sheet.getRange(2, 7, lastRow - 1, 2).setHorizontalAlignment('left');

    const generoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(FEST_GENRE_DROPDOWN, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, 2, lastRow - 1, 1).setDataValidation(generoRule);
  }
`,
'dropdown genero'
);

mustReplace(
"  const headerRange = sheet.getRange(1, 1, 1, lastCol);",
"  const tax = parseSheetTaxonomy_(sheet.getName());\n  const tabColors = { URBAN: '#FB8C00', POP: '#EC407A', INDIE: '#546E7A', ROCK: '#E53935', ELECTR: '#00ACC1', JAZZ: '#3949AB', FLAM: '#D81B60', RUMBA: '#43A047', MEC: '#6D4C41', MFR: '#8D6E63', PTE: '#757575' };\n  if (tabColors[tax.genre]) {\n    try { sheet.setTabColor(tabColors[tax.genre]); } catch (err) {}\n  }\n\n  const headerRange = sheet.getRange(1, 1, 1, lastCol);",
'tab color'
);

const insertPoint = "function mostrarGuiaIntegracionCRMFestivales() {";
if (!text.includes('function auditarClasificacionGeneroTamanoCRMFestivales() {')) {
  const block = `function auditarClasificacionGeneroTamanoCRMFestivales() {\n  if (!validarSesionSegura_('Auditar genero + tamano S/L/XL')) return;\n\n  const ss = SpreadsheetApp.getActiveSpreadsheet();\n  const sheets = getFestivalSheets_(ss);\n  if (!sheets.length) {\n    SpreadsheetApp.getUi().alert('No encontre pestanas de festivales para auditar clasificacion.');\n    return;\n  }\n\n  let totalRows = 0;\n  let noGenero = 0;\n  let noAforo = 0;\n  let mismatchGenero = 0;\n  let mismatchTamano = 0;\n  const examples = [];\n\n  sheets.forEach((sheet) => {\n    const values = sheet.getDataRange().getValues();\n    if (values.length < 2) return;\n\n    const map = buildHeaderMap_(values[0]);\n    const tax = parseSheetTaxonomy_(sheet.getName());\n\n    for (let r = 1; r < values.length; r++) {\n      const row = values[r];\n      const nombre = cleanText_(valueAt_(row, map.nombre));\n      if (!nombre) continue;\n\n      totalRows++;\n      const generoRow = normalizeGenreCode_(valueAt_(row, map.genero));\n      const sizeRow = sizeCodeFromAforo_(valueAt_(row, map.aforo));\n\n      if (!generoRow) noGenero++;\n      if (!sizeRow) noAforo++;\n\n      if (tax.genre && tax.genre !== 'PTE' && generoRow && generoRow !== tax.genre) {\n        mismatchGenero++;\n        if (examples.length < 18) examples.push(sheet.getName() + '!A' + (r + 1) + ' -> genero=' + generoRow + ' (esperado ' + tax.genre + ')');\n      }\n\n      if (tax.size && sizeRow && sizeRow !== tax.size) {\n        mismatchTamano++;\n        if (examples.length < 18) examples.push(sheet.getName() + '!A' + (r + 1) + ' -> tamano=' + sizeRow + ' (esperado ' + tax.size + ')');\n      }\n    }\n  });\n\n  const resumen = [\n    '🧭 Auditoria de clasificacion (' + FEST_ARCHITECT + ')',\n    '',\n    'Filas revisadas: ' + totalRows,\n    'Sin genero: ' + noGenero,\n    'Sin aforo/tamano: ' + noAforo,\n    'Desajustes de genero: ' + mismatchGenero,\n    'Desajustes de tamano: ' + mismatchTamano,\n    '',\n    'Reglas de tamano:',\n    '- S: 0 a 1000 (incluido 1000)',\n    '- L: 1001 a 9999',\n    '- XL: 10000 en adelante',\n    '',\n    examples.length ? 'Muestras:\\n' + examples.join('\\n') : 'No se detectaron desajustes en la muestra analizada.'\n  ].join('\\n');\n\n  SpreadsheetApp.getUi().alert(resumen);\n}\n\n`;
  text = text.replace(insertPoint, block + insertPoint);
}

fs.writeFileSync(path, text, 'utf8');
console.log('OK');
