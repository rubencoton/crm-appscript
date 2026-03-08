const fs = require('fs');
const vm = require('vm');

const file = 'C:/Users/elrub/Desktop/CARPETA CODEX/01_PROYECTOS/festivales-github/CRM_FESTIVALES_ENGINE.gs';
const code = fs.readFileSync(file, 'utf8');

const fnMatches = [...code.matchAll(/function\s+([A-Za-z0-9_]+)\s*\(/g)].map((m) => m[1]);
const fnCount = fnMatches.reduce((acc, n) => ((acc[n] = (acc[n] || 0) + 1), acc), {});
const duplicated = Object.entries(fnCount).filter(([, c]) => c > 1);

const context = {
  console,
  Utilities: { sleep: () => {} },
  UrlFetchApp: { fetch: () => { throw new Error('NETWORK_DISABLED_IN_TEST'); } },
  CacheService: { getUserCache: () => ({ put: () => {}, get: () => null }) },
  SpreadsheetApp: {},
  ScriptApp: {},
  HtmlService: {},
  MailApp: {}
};
vm.createContext(context);
vm.runInContext(code, context, { filename: file });

const failures = [];
const notes = [];
function assertEq(actual, expected, label) {
  if (actual !== expected) failures.push(`${label} | esperado=${JSON.stringify(expected)} actual=${JSON.stringify(actual)}`);
}
function assertTrue(cond, label) {
  if (!cond) failures.push(label);
}

// Tests deterministas de borde
assertEq(context.sizeCodeFromAforo_('0'), 'S', 'sizeCodeFromAforo_ 0');
assertEq(context.sizeCodeFromAforo_('1000'), 'S', 'sizeCodeFromAforo_ 1000');
assertEq(context.sizeCodeFromAforo_('1001'), 'L', 'sizeCodeFromAforo_ 1001');
assertEq(context.sizeCodeFromAforo_('9999'), 'L', 'sizeCodeFromAforo_ 9999');
assertEq(context.sizeCodeFromAforo_('10000'), 'XL', 'sizeCodeFromAforo_ 10000');
assertEq(context.formatSpanishPhone_('612345678'), '+34 612 345 678', 'formatSpanishPhone_ local');
assertEq(context.formatSpanishPhone_('+34 612345678'), '+34 612 345 678', 'formatSpanishPhone_ +34');
assertEq(context.formatSpanishPhone_('12345'), '', 'formatSpanishPhone_ invalido');
assertTrue(context.isValidEmailList_('a@b.com; c@d.es'), 'isValidEmailList_ lista valida');
assertTrue(!context.isValidEmailList_('a@b,com'), 'isValidEmailList_ invalido');
assertEq(context.normalizeGenreCode_('🎤 POP'), 'POP', 'normalizeGenreCode_ emoji pop');
assertEq(context.normalizeGenreCode_('MUSICA REGIONAL'), 'MFR', 'normalizeGenreCode_ regional');

// Sanidad de menu -> handlers existentes
const addItems = [...code.matchAll(/\.addItem\([^,]+,\s*'([^']+)'\)/g)].map((m) => m[1]);
const missingHandlers = addItems.filter((h) => typeof context[h] !== 'function');
if (missingHandlers.length) {
  failures.push('Handlers de menu inexistentes: ' + missingHandlers.join(', '));
}

if (duplicated.length) {
  failures.push('Funciones duplicadas detectadas: ' + duplicated.map(([n, c]) => `${n}x${c}`).join(', '));
}

// Robustez parser Gemini
assertEq(context.parseGeminiJson_(''), null, 'parseGeminiJson_ vacio');
assertEq(context.parseGeminiJson_('{"x":1}'), null, 'parseGeminiJson_ sin candidates');
assertEq(context.parseGeminiJson_('{mal json'), null, 'parseGeminiJson_ mal json');

// Fuzz agresivo sobre helpers puros
function randInt(max) { return Math.floor(Math.random() * max); }
function randStr() {
  const samples = ['abc', '  ', 'ñáéíóú', '🎤 POP', 'URBAN', 'MFR', '++--', '1.000', 'a@b.com', 'sin informacion', 'NULL', '0034612345678'];
  const base = samples[randInt(samples.length)];
  return base + (Math.random() < 0.3 ? String(randInt(99999)) : '');
}
function randVal() {
  const t = randInt(9);
  if (t === 0) return null;
  if (t === 1) return undefined;
  if (t === 2) return randInt(200000) - 1000;
  if (t === 3) return randStr();
  if (t === 4) return { x: randStr() };
  if (t === 5) return ['a', randStr()];
  if (t === 6) return true;
  if (t === 7) return false;
  return '';
}

const fuzzFns = [
  'cleanText_', 'normalizeHeader_', 'parseAforo_', 'normalizeAforoForDisplay_',
  'normalizeEmailCell_', 'normalizePhoneCell_', 'formatSpanishPhone_',
  'normalizeContactName_', 'isValidEmailList_', 'isPlaceholderText_',
  'normalizeGenreCode_', 'sizeCodeFromAforo_'
];

const fuzzErrors = {};
for (const fn of fuzzFns) fuzzErrors[fn] = 0;

for (let i = 0; i < 12000; i++) {
  const v = randVal();
  for (const fn of fuzzFns) {
    try {
      context[fn](v);
    } catch (e) {
      fuzzErrors[fn]++;
    }
  }
}

const badFuzz = Object.entries(fuzzErrors).filter(([, n]) => n > 0);
if (badFuzz.length) {
  failures.push('Errores en fuzz: ' + badFuzz.map(([n, c]) => `${n}=${c}`).join(', '));
} else {
  notes.push('Fuzz de 12k iteraciones sin throws en helpers puros.');
}

console.log('--- AUDIT SUMMARY ---');
console.log('Funciones totales:', fnMatches.length);
console.log('Handlers de menu:', addItems.length);
notes.forEach((n) => console.log('NOTE:', n));
if (!failures.length) {
  console.log('RESULT: OK (sin fallos en tests locales de robustez).');
} else {
  console.log('RESULT: FAIL (' + failures.length + ' hallazgos)');
  failures.forEach((f, i) => console.log(`${i + 1}. ${f}`));
  process.exitCode = 2;
}

