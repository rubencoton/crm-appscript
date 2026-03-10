#!/usr/bin/env node
/* eslint-disable no-console */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const GIT = findGit();

const result = { checks: 0, pass: 0, fail: 0, warnings: [] };

main();

function main() {
  log('AUDITORIA AUTOMATICA');
  log('Repo: ' + ROOT);

  section('Config', checkConfig);
  section('Syntax', checkSyntax);
  section('Secrets', checkSecrets);
  section('Stress CRM ayudas', stressAyudas);
  section('Stress CRM festivales', stressFestivales);

  console.log('\nRESUMEN');
  console.log('Checks:', result.checks);
  console.log('Pass:', result.pass);
  console.log('Fail:', result.fail);
  if (result.warnings.length) {
    console.log('Warnings:');
    result.warnings.forEach((w) => console.log('-', w));
  }

  if (result.fail > 0) process.exit(1);
}

function section(name, fn) {
  console.log('\n=== ' + name + ' ===');
  try {
    fn();
  } catch (e) {
    fail('SECTION ' + name, e.message || String(e));
  }
}

function check(name, fn) {
  result.checks++;
  try {
    fn();
    result.pass++;
    console.log('[PASS]', name);
  } catch (e) {
    fail(name, e.message || String(e));
  }
}

function fail(name, msg) {
  result.fail++;
  console.log('[FAIL]', name, '=>', msg);
}

function warn(msg) {
  result.warnings.push(msg);
  console.log('[WARN]', msg);
}

function checkConfig() {
  const claspIgnore = read(path.join(ROOT, 'crm-ayudas-subvenciones', '.claspignore'));
  const gitIgnore = read(path.join(ROOT, '.gitignore'));

  check('.claspignore expected entries', () => {
    assert(claspIgnore.includes('!appsscript.json'), 'falta !appsscript.json');
    assert(claspIgnore.includes('!Code.js'), 'falta !Code.js');
    assert(claspIgnore.includes('!DecisionInstantanea.gs'), 'falta !DecisionInstantanea.gs');
  });

  check('.gitignore blocks DecisionInstantanea.js', () => {
    assert(gitIgnore.includes('crm-ayudas-subvenciones/DecisionInstantanea.js'), 'falta regla ignore');
  });
}

function checkSyntax() {
  const files = gitLs(['*.js', '*.gs']);
  check('tracked js/gs not empty', () => assert(files.length > 0, 'sin archivos js/gs'));

  files.forEach((f) => {
    check('syntax ' + f, () => {
      // eslint-disable-next-line no-new-func
      new Function(read(path.join(ROOT, f)));
    });
  });
}

function checkSecrets() {
  const files = gitLs([]);
  const hits = [];

  files.forEach((f) => {
    const txt = readMaybe(path.join(ROOT, f));
    if (!txt) return;
    const lines = txt.split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const ln = i + 1;

      if (/AIza[0-9A-Za-z\-_]{20,}/.test(line)) {
        if (!line.includes('placeholder="AIza')) hits.push(`${f}:${ln} api key real`);
      }
      if (line.includes('rubencoton26')) hits.push(`${f}:${ln} password historica`);

      const m = line.match(/(?:const|let|var)\s+[A-Za-z0-9_]*(?:API_KEY|PASSWORD)[A-Za-z0-9_]*\s*=\s*['"`]([^'"`]+)['"`]/);
      if (m && m[1]) {
        const v = String(m[1]).trim();
        if (!['REDACTED_API_KEY', 'REDACTED_PASSWORD', 'PLACEHOLDER_API_KEY', 'PLACEHOLDER_PASSWORD'].includes(v)) {
          hits.push(`${f}:${ln} posible secreto hardcodeado`);
        }
      }
    }
  });

  check('no leaked secrets', () => {
    assert(hits.length === 0, hits.slice(0, 20).join('\n'));
  });
}

function stressAyudas() {
  const src = read(path.join(ROOT, 'crm-ayudas-subvenciones', 'Code.js')) + '\n' + read(path.join(ROOT, 'crm-ayudas-subvenciones', 'DecisionInstantanea.gs'));
  const ctx = makeGasContext();
  vm.createContext(ctx);
  vm.runInContext(src, ctx, { timeout: 15000 });

  check('ayudas invalid day rejected', () => assert(ctx.parseFechaLimite_('32/01/2026') === null, 'acepto dia invalido'));
  check('ayudas invalid month rejected', () => assert(ctx.parseFechaLimite_('10/13/2026') === null, 'acepto mes invalido'));
  check('ayudas json fenced', () => assert(ctx.extractJsonBlock_('```json\n{"a":1}\n```') === '{"a":1}', 'fallo json fenced'));
  check('ayudas scenario decision returns', () => assert(!!ctx.construirDecisionEscenarioDecision_(5, 6, 2).decision, 'sin decision'));

  check('ayudas date fuzz no coercion', () => {
    for (let i = 0; i < 240; i++) {
      const d = rand(0, 39);
      const m = rand(0, 14);
      const y = rand(2020, 2031);
      const s = p2(d) + '/' + p2(m) + '/' + y;
      const out = ctx.parseFechaLimite_(s);
      if (!out) continue;
      assert(out.date.getDate() === d && out.date.getMonth() + 1 === m && out.date.getFullYear() === y, 'coercion ' + s);
    }
  });
}

function stressFestivales() {
  const src = read(path.join(ROOT, 'CRM_FESTIVALES_ENGINE.js'));
  const ctx = makeGasContext({ FEST_GEMINI_API_KEY: 'DUMMY', FEST_SECURITY_PASSWORD: 'dummy-pass-123' });

  let fetchCalls = 0;
  ctx.UrlFetchApp = {
    fetch: (url) => {
      fetchCalls++;
      if (url.includes('/models?key=')) {
        return {
          getResponseCode: () => 200,
          getContentText: () => JSON.stringify({ models: [{ name: 'models/gemini-3.1-pro-preview' }, { name: 'models/gemini-2.5-image' }] })
        };
      }
      return { getResponseCode: () => 500, getContentText: () => '{"error":{"message":"x"}}' };
    }
  };

  vm.createContext(ctx);
  vm.runInContext(src, ctx, { timeout: 15000 });

  check('fest parse embedded json', () => {
    const raw = JSON.stringify({ candidates: [{ content: { parts: [{ text: 'antes {"a":1} despues' }] } }] });
    const out = ctx.parseGeminiJson_(raw);
    assert(out && out.a === 1, 'no extrajo json incrustado');
  });

  check('fest parse fenced json', () => {
    const raw = JSON.stringify({ candidates: [{ content: { parts: [{ text: '```json\n{"a":1}\n```' }] } }] });
    const out = ctx.parseGeminiJson_(raw);
    assert(out && out.a === 1, 'no parseo fenced');
  });

  check('fest model filter works', () => {
    assert(ctx.isGeminiTextModelCandidate_('gemini-2.5-image') === false, 'no filtro image');
    assert(ctx.isGeminiTextModelCandidate_('gemini-3.1-pro-preview') === true, 'filtro modelo valido');
  });

  check('fest model discovery with fetch stub', () => {
    const models = ctx.getGeminiModelCandidates_();
    assert(Array.isArray(models) && models.length > 0, 'sin modelos');
  });

  check('fest fuzz parser no throw', () => {
    for (let i = 0; i < 260; i++) {
      const t = Math.random().toString(36).repeat(rand(1, 6));
      const payload = Math.random() < 0.5 ? t : JSON.stringify({ candidates: [{ content: { parts: [{ text: t }] } }] });
      const out = ctx.parseGeminiJson_(payload);
      if (out !== null) assert(typeof out === 'object', 'salida no valida');
    }
  });

  if (fetchCalls === 0) warn('No hubo llamadas fetch en stress festivales');
}

function makeGasContext(initialProps) {
  const props = new Map(Object.entries(initialProps || {}).map(([k, v]) => [k, String(v)]));
  const cache = new Map();

  return {
    console,
    Math,
    Date,
    JSON,
    SpreadsheetApp: {
      BorderStyle: { SOLID: 'SOLID' },
      ProtectionType: { RANGE: 'RANGE', SHEET: 'SHEET' },
      newDataValidation: () => ({ requireValueInList() { return this; }, setAllowInvalid() { return this; }, build() { return {}; } }),
      newConditionalFormatRule: () => ({ whenFormulaSatisfied() { return this; }, setBackground() { return this; }, setRanges() { return this; }, build() { return {}; } }),
      getUi: () => ({ alert() {}, createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() { return this; } }), prompt: () => ({ getSelectedButton: () => 'CANCEL', getResponseText: () => '' }), ButtonSet: { OK_CANCEL: 'OK_CANCEL' }, Button: { OK: 'OK' } }),
      getActive: () => ({ getSpreadsheetTimeZone: () => 'Europe/Madrid' }),
      getActiveSpreadsheet: () => ({ getSheets: () => [] })
    },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => ({ forSpreadsheet() { return this; }, onOpen() { return this; }, timeBased() { return this; }, after() { return this; }, create() {} }), getScriptId: () => 'audit-script-id' },
    LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }), getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (props.has(k) ? props.get(k) : ''), setProperty: (k, v) => props.set(k, String(v)), setProperties: (o) => Object.keys(o || {}).forEach((k) => props.set(k, String(o[k]))), deleteProperty: (k) => props.delete(k), deleteAllProperties: () => props.clear() }) },
    CacheService: { getScriptCache: () => ({ get: (k) => (cache.has(k) ? cache.get(k) : null), put: (k, v) => cache.set(k, String(v)), remove: (k) => cache.delete(k) }), getUserCache: () => ({ put() {}, get: () => '' }) },
    Session: { getScriptTimeZone: () => 'Europe/Madrid', getEffectiveUser: () => ({ getEmail: () => 'audit@example.com' }) },
    HtmlService: { createHtmlOutput: () => ({ setWidth() { return this; }, setHeight() { return this; } }) },
    UrlFetchApp: { fetch: () => ({ getResponseCode: () => 500, getContentText: () => '{}' }) },
    DriveApp: {},
    Utilities: {
      formatDate: (d, tz, fmt) => {
        const dd = p2(d.getDate()); const mm = p2(d.getMonth() + 1); const yyyy = d.getFullYear();
        const hh = p2(d.getHours()); const mi = p2(d.getMinutes()); const ss = p2(d.getSeconds());
        if (fmt === 'dd/MM/yyyy') return `${dd}/${mm}/${yyyy}`;
        if (fmt === 'HH:mm:ss') return `${hh}:${mi}:${ss}`;
        if (fmt === 'dd/MM/yyyy HH:mm:ss') return `${dd}/${mm}/${yyyy} ${hh}:${mi}:${ss}`;
        return d.toISOString();
      },
      sleep() {},
      base64Encode: () => '',
      computeDigest: () => [1, 2, 3, 4],
      DigestAlgorithm: { SHA_256: 'SHA_256' }
    }
  };
}

function gitLs(patterns) {
  const args = ['ls-files'].concat(patterns || []);
  const raw = execFileSync(GIT, args, { cwd: ROOT, encoding: 'utf8' });
  return String(raw || '').split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
}

function findGit() {
  const bins = ['git', 'C:\\Program Files\\Git\\cmd\\git.exe', 'C:\\Program Files\\Git\\bin\\git.exe'];
  for (const b of bins) {
    try { execFileSync(b, ['--version'], { stdio: 'ignore' }); return b; } catch (e) {}
  }
  throw new Error('git no encontrado');
}

function read(p) { return fs.readFileSync(p, 'utf8'); }
function readMaybe(p) { try { return fs.readFileSync(p, 'utf8'); } catch (e) { return null; } }
function assert(cond, msg) { if (!cond) throw new Error(msg || 'assert'); }
function rand(min, max) { return Math.floor(Math.random() * (max - min)) + min; }
function p2(n) { return String(n).padStart(2, '0'); }
function log(msg) { console.log('[INFO]', msg); }
