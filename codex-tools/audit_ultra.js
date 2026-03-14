#!/usr/bin/env node
/* eslint-disable no-console */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const https = require('https');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const GIT = findGit();
const LIVE_RETRYABLE_STATUS = { 0: true, 408: true, 409: true, 425: true, 429: true, 500: true, 502: true, 503: true, 504: true };
const LIVE_MAX_RETRIES = 4;
const LIVE_BASE_DELAY_MS = 600;

const result = { checks: 0, pass: 0, fail: 0, warnings: [] };

main().catch((err) => {
  fail('ULTRA MAIN', err && err.message ? err.message : String(err));
  finish();
  process.exit(1);
});

async function main() {
  log('AUDITORIA ULTRA PROFUNDA');
  log('Repo: ' + ROOT);

  await section('Baseline Stability', runAuditAutoStability);
  await section('Deep Stress Local', deepStressLocal);
  await section('Prompt Integrity', promptIntegrity);
  await section('Secrets Deep Scan', secretsDeepScan);
  await section('Live API Validation', liveApiValidation);

  finish();
  if (result.fail > 0) process.exit(1);
}

function finish() {
  console.log('\nRESUMEN ULTRA');
  console.log('Checks:', result.checks);
  console.log('Pass:', result.pass);
  console.log('Fail:', result.fail);
  if (result.warnings.length) {
    console.log('Warnings:');
    result.warnings.forEach((w) => console.log('-', w));
  }
}

async function section(name, fn) {
  console.log('\n=== ' + name + ' ===');
  try {
    await fn();
  } catch (e) {
    fail('SECTION ' + name, e && e.message ? e.message : String(e));
  }
}

function check(name, fn) {
  result.checks++;
  try {
    fn();
    result.pass++;
    console.log('[PASS]', name);
  } catch (e) {
    fail(name, e && e.message ? e.message : String(e));
  }
}

async function checkAsync(name, fn) {
  result.checks++;
  try {
    await fn();
    result.pass++;
    console.log('[PASS]', name);
  } catch (e) {
    fail(name, e && e.message ? e.message : String(e));
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

function log(msg) {
  console.log('[INFO]', msg);
}

async function runAuditAutoStability() {
  for (let i = 1; i <= 3; i++) {
    check('audit:auto stable run #' + i, () => {
      const out = execFileSync(process.execPath, [path.join(__dirname, 'audit_auto.js')], {
        cwd: ROOT,
        encoding: 'utf8',
        maxBuffer: 30 * 1024 * 1024
      });
      assert(String(out).indexOf('Fail: 0') !== -1, 'audit:auto no quedo en 0 fallos');
    });
  }
}

async function deepStressLocal() {
  const ayudasSrc = read(path.join(ROOT, 'crm-ayudas-subvenciones', 'Code.js')) + '\n' + read(path.join(ROOT, 'crm-ayudas-subvenciones', 'DecisionInstantanea.gs'));
  const festJsSrc = read(path.join(ROOT, 'CRM_FESTIVALES_ENGINE.js'));
  const festGsSrc = read(path.join(ROOT, 'CRM_FESTIVALES_ENGINE.gs'));

  const ayudasCtx = makeGasContext();
  vm.createContext(ayudasCtx);
  vm.runInContext(ayudasSrc, ayudasCtx, { timeout: 20000 });

  check('ayudas parseFecha fuzz 5000 no coercion', () => {
    for (let i = 0; i < 5000; i++) {
      const d = rand(0, 40);
      const m = rand(0, 15);
      const y = rand(2020, 2035);
      const s = p2(d) + '/' + p2(m) + '/' + y;
      const out = ayudasCtx.parseFechaLimite_(s);
      if (!out) continue;
      assert(out.date.getDate() === d, 'dia coercionado: ' + s);
      assert(out.date.getMonth() + 1 === m, 'mes coercionado: ' + s);
      assert(out.date.getFullYear() === y, 'anio coercionado: ' + s);
    }
  });

  check('ayudas estado inscripcion coincide referencia 1000 casos', () => {
    function subMonthsSafe(dateObj, months) {
      const base = new Date(dateObj);
      const day = base.getDate();
      const first = new Date(base.getFullYear(), base.getMonth() - months, 1, 12, 0, 0, 0);
      const lastDay = new Date(first.getFullYear(), first.getMonth() + 1, 0).getDate();
      const safeDay = Math.min(day, lastDay);
      return new Date(first.getFullYear(), first.getMonth(), safeDay, 12, 0, 0, 0);
    }

    function expected(fecha) {
      const lim = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate(), 23, 59, 59, 999);
      const ini = subMonthsSafe(lim, 3);
      ini.setHours(0, 0, 0, 0);
      const hoy = new Date();
      hoy.setHours(0, 0, 0, 0);
      if (hoy > lim) return 'CERRADA';
      if (hoy >= ini && hoy <= lim) return 'ABIERTA';
      return 'SIN PUBLICAR';
    }

    for (let i = 0; i < 1000; i++) {
      const dt = new Date();
      dt.setDate(dt.getDate() + rand(-365, 366));
      const s = p2(dt.getDate()) + '/' + p2(dt.getMonth() + 1) + '/' + dt.getFullYear();
      const got = ayudasCtx.estadoInscripcionDesdeFecha_(s);
      const exp = expected(dt);
      assert(got === exp, 'estado distinto para ' + s + ' => got=' + got + ' exp=' + exp);
    }
  });

  check('ayudas normalizarObjeto invariantes 1200 casos', () => {
    const allowedIns = { ABIERTA: true, CERRADA: true, 'SIN PUBLICAR': true };
    const allowedPremio = { ECONOMICO: true, SERVICIO: true, ACTUACION: true, RESIDENCIA: true, VARIOS: true };
    const req = [
      'nombreConcurso', '_razonamiento_logico', 'inscripcion', 'fechaLimite', 'fechasDesarrollo',
      'tipoPremio', 'detallePremio', 'dirigidoA', 'municipio', 'provincia', 'pais',
      'link1', 'link2', 'link3', 'email', 'telefono', 'notas'
    ];

    for (let i = 0; i < 1200; i++) {
      const raw = {};
      if (Math.random() < 0.4) raw.nombreConcurso = 'Concurso ' + i;
      if (Math.random() < 0.5) raw.inscripcion = ['ABIERTA', 'CERRADA', 'SIN PUBLICAR', '', 'abierta'][rand(0, 5)];
      if (Math.random() < 0.5) raw.fechaLimite = [
        '15/12/2030', 'ESTIMADO: 01/02/2031', '32/01/2030', '', null
      ][rand(0, 5)];
      if (Math.random() < 0.5) raw.tipoPremio = ['ECONOMICO', 'SERVICIO', 'ACTUACION', 'RESIDENCIA', 'VARIOS', 'rare'][rand(0, 6)];
      if (Math.random() < 0.5) raw.link1 = ['https://x.com', 'nota', '', null][rand(0, 4)];
      if (Math.random() < 0.5) raw._razonamiento_logico = Math.random() < 0.5 ? 'Texto breve' : 'Razonamiento extenso con evidencia.';

      const out = ayudasCtx.normalizarObjetoIA_(raw);
      assert(out && typeof out === 'object', 'normalizarObjetoIA_ devolvio null');
      assert(allowedIns[String(out.inscripcion)], 'inscripcion invalida: ' + out.inscripcion);
      assert(allowedPremio[String(out.tipoPremio)], 'tipo premio invalido: ' + out.tipoPremio);

      for (let k = 0; k < req.length; k++) {
        const key = req[k];
        assert(!!String(out[key] || '').trim(), 'campo vacio: ' + key);
      }
    }
  });

  check('ayudas llamarIA retry sin rotacion (503->200)', () => {
    const props = ayudasCtx.PropertiesService.getScriptProperties();
    props.setProperty('GEMINI_API_KEY', 'dummy');
    props.setProperty('IDX_MODEL', '0');
    props.deleteProperty('STOP_REQUESTED');

    try { ayudasCtx.CacheService.getScriptCache().remove('CRM_GEM_MODELS_V2'); } catch (e) {}

    const prev = ayudasCtx.UrlFetchApp.fetch;
    let modelCalls = 0;
    let genCalls = 0;

    try {
      ayudasCtx.UrlFetchApp.fetch = (url) => {
        const u = String(url || '');
        if (u.indexOf('/v1beta/models?key=') !== -1) {
          modelCalls++;
          return {
            getResponseCode: () => 200,
            getContentText: () => JSON.stringify({ models: [{ name: 'models/gemini-3.1-pro-preview' }] })
          };
        }

        genCalls++;
        if (genCalls === 1) {
          return {
            getResponseCode: () => 503,
            getContentText: () => JSON.stringify({ error: { message: 'overloaded' } })
          };
        }

        return {
          getResponseCode: () => 200,
          getContentText: () => JSON.stringify({
            candidates: [{ content: { parts: [{ text: '{"nombreConcurso":"C","_razonamiento_logico":"Confianza: ALTA.","inscripcion":"ABIERTA","fechaLimite":"15/12/2030","fechasDesarrollo":"Diciembre","tipoPremio":"ECONOMICO","detallePremio":"100","dirigidoA":"Bandas","municipio":"Madrid","provincia":"Madrid","pais":"Espana","link1":"https://a.com","link2":"","link3":"","email":"a@b.com","telefono":"+34 600 000 000","notas":"ok"}' }] } }]
          })
        };
      };

      const out = ayudasCtx.llamarIA_('PROMPT_X', { type: 'none' }, false);
      assert(out && out.nombreConcurso === 'C', 'sin salida valida en retry');
      assert(modelCalls >= 1, 'no consulto modelos');
      assert(genCalls >= 2, 'no reintento tras 503');
      const idxStored = Number(props.getProperty('IDX_MODEL') || '-1');
      assert(idxStored >= 0, 'IDX_MODEL no persistido tras retry');
    } finally {
      ayudasCtx.UrlFetchApp.fetch = prev;
    }
  });

  check('fest parser js vs gs consistencia 2500 fuzz', () => {
    const jsCtx = makeGasContext({ FEST_GEMINI_API_KEY: 'DUMMY', FEST_SECURITY_PASSWORD: 'x' });
    const gsCtx = makeGasContext({ FEST_GEMINI_API_KEY: 'DUMMY', FEST_SECURITY_PASSWORD: 'x' });
    vm.createContext(jsCtx);
    vm.createContext(gsCtx);
    vm.runInContext(festJsSrc, jsCtx, { timeout: 20000 });
    vm.runInContext(festGsSrc, gsCtx, { timeout: 20000 });

    for (let i = 0; i < 2500; i++) {
      const mode = rand(0, 4);
      let payload = '';
      if (mode === 0) {
        payload = JSON.stringify({ candidates: [{ content: { parts: [{ text: '{"a":1}' }] } }] });
      } else if (mode === 1) {
        payload = JSON.stringify({ candidates: [{ content: { parts: [{ text: '```json\\n{"a":1}\\n```' }] } }] });
      } else if (mode === 2) {
        payload = JSON.stringify({ candidates: [{ content: { parts: [{ text: 'abc {"a":1} def' }] } }] });
      } else {
        const r = Math.random().toString(36).repeat(rand(1, 6));
        payload = Math.random() < 0.5 ? r : JSON.stringify({ candidates: [{ content: { parts: [{ text: r }] } }] });
      }

      const a = jsCtx.parseGeminiJson_(payload);
      const b = gsCtx.parseGeminiJson_(payload);

      if (a === null && b === null) continue;
      assert((a === null) === (b === null), 'js/gs inconsistente en null');
      if (a && b) {
        assert(typeof a === 'object' && typeof b === 'object', 'tipo no objeto en js/gs');
      }
    }
  });
}

async function promptIntegrity() {
  const ayudas = read(path.join(ROOT, 'crm-ayudas-subvenciones', 'Code.js'));
  const festJs = read(path.join(ROOT, 'CRM_FESTIVALES_ENGINE.js'));
  const festGs = read(path.join(ROOT, 'CRM_FESTIVALES_ENGINE.gs'));

  check('prompt ayudas mantiene reglas criticas', () => {
    assert(ayudas.includes('fechaLimite es el dato mas importante'), 'falta regla fechaLimite');
    assert(ayudas.includes('Inscripcion ABIERTA cuando hoy esta entre (fechaLimite - 3 meses) y fechaLimite, inclusive.'), 'falta regla ventana 3 meses');
    assert(ayudas.includes('Aplica razonamiento profundo'), 'falta razonamiento profundo');
    assert(ayudas.includes('responseMimeType'), 'falta responseMimeType');
    assert(ayudas.includes('required: AI_REQUIRED_FIELDS.slice()'), 'falta required fields');
  });

  check('prompt festivales js/gs mantiene no invencion', () => {
    assert(festJs.includes('Nunca inventes datos no presentes en la entrada.'), 'js: falta no inventar');
    assert(festJs.includes('Responde SOLO en JSON valido, sin markdown.'), 'js: falta json estricto');
    assert(festGs.includes('Nunca inventes datos no presentes en la entrada.'), 'gs: falta no inventar');
    assert(festGs.includes('Responde SOLO en JSON valido, sin markdown.'), 'gs: falta json estricto');
  });
}

async function secretsDeepScan() {
  const files = gitLs([]);
  const hits = [];

  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    if (f.indexOf('node_modules/') === 0) continue;
    if (f.indexOf('_full_workspace_backup/') === 0) continue;

    const txt = readMaybe(path.join(ROOT, f));
    if (!txt) continue;
    const lines = txt.split(/\r?\n/);

    for (let ln = 0; ln < lines.length; ln++) {
      const line = lines[ln];
      const n = ln + 1;

      if (/AIza[0-9A-Za-z\-_]{20,}/.test(line)) hits.push(f + ':' + n + ' api key visible');
      if (/sk-[A-Za-z0-9]{20,}/.test(line)) hits.push(f + ':' + n + ' token pattern visible');
      if (/(password|passwd|secret)\s*[:=]\s*['"][^'"]+['"]/i.test(line)) {
        if (!/REDACTED|PLACEHOLDER|dummy|DUMMY|''/.test(line)) hits.push(f + ':' + n + ' potential secret literal');
      }
    }
  }

  check('no leaked secrets in tracked sources', () => {
    assert(hits.length === 0, hits.slice(0, 20).join('\n'));
  });
}

async function liveApiValidation() {
  const key = String(process.env.GEMINI_API_KEY || '').trim();
  check('GEMINI_API_KEY provided for live validation', () => {
    assert(!!key, 'falta variable de entorno GEMINI_API_KEY');
  });
  if (!key) return;

  const list = await httpReqWithRetry('GET', 'https://generativelanguage.googleapis.com/v1beta/models?key=' + encodeURIComponent(key));
  await checkAsync('live list models returns HTTP 200', async () => {
    assert(list.status === 200, 'HTTP ' + list.status + ' | ' + truncate(String(list.body), 240));
  });
  if (list.status !== 200) return;

  let catalog = [];
  try {
    catalog = (JSON.parse(list.body).models || []).map((m) => String(m.name || '').replace(/^models\//, ''));
  } catch (e) {
    catalog = [];
  }

  check('live catalog not empty', () => {
    assert(Array.isArray(catalog) && catalog.length > 0, 'catalogo vacio');
  });

  const modelCandidates = extractGeminiCandidatesFromSource();
  const preferred = prioritizeModels(modelCandidates);
  const present = preferred.filter((m) => catalog.indexOf(m) !== -1);

  check('critical models present in live catalog', () => {
    const must = ['gemini-3.1-pro-preview', 'gemini-2.5-pro', 'gemini-2.5-flash'];
    for (let i = 0; i < must.length; i++) {
      assert(catalog.indexOf(must[i]) !== -1, 'falta modelo critico: ' + must[i]);
    }
  });

  const targets = present.slice(0, 6);
  check('have at least 3 present target models', () => {
    assert(targets.length >= 3, 'solo hay ' + targets.length + ' modelos target presentes');
  });

  for (let i = 0; i < targets.length; i++) {
    const model = targets[i];

    await checkAsync('live gen simple OK [' + model + ']', async () => {
      const body = {
        contents: [{ role: 'user', parts: [{ text: 'Responde solo OK' }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 2048 }
      };
      const res = await httpReqWithRetry('POST', 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + encodeURIComponent(key), body);
      assert(res.status === 200, 'HTTP ' + res.status + ' | ' + truncate(String(res.body), 220));
      const txt = extractCandidateTextFromRaw(res.body);
      assert(!!txt, 'respuesta vacia');
    });

    await checkAsync('live gen json strict [' + model + ']', async () => {
      const body = {
        systemInstruction: { parts: [{ text: 'Devuelve solo JSON valido.' }] },
        contents: [{ role: 'user', parts: [{ text: 'Devuelve {"ok":true,"source":"live"}' }] }],
        generationConfig: {
          responseMimeType: 'application/json',
          temperature: 0.1,
          maxOutputTokens: 2048
        }
      };
      const res = await httpReqWithRetry('POST', 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + encodeURIComponent(key), body);
      assert(res.status === 200, 'HTTP ' + res.status + ' | ' + truncate(String(res.body), 220));
      const txt = extractCandidateTextFromRaw(res.body);
      assert(!!txt, 'respuesta json vacia');
    });
  }
}

function extractGeminiCandidatesFromSource() {
  const files = [
    path.join(ROOT, 'crm-ayudas-subvenciones', 'Code.js'),
    path.join(ROOT, 'CRM_FESTIVALES_ENGINE.js'),
    path.join(ROOT, 'CRM_FESTIVALES_ENGINE.gs')
  ];

  const out = [];
  const seen = {};
  for (let i = 0; i < files.length; i++) {
    const txt = readMaybe(files[i]) || '';
    const matches = txt.match(/gemini[-a-z0-9\.]+/gi) || [];
    for (let m = 0; m < matches.length; m++) {
      const model = String(matches[m]).toLowerCase();
      if (model.indexOf('gemini') !== 0) continue;
      if (/(image|audio|tts|embedding|aqa|veo|imagen|gemma|robotics|computer-use|deep-research)/i.test(model)) continue;
      if (!seen[model]) {
        seen[model] = true;
        out.push(model);
      }
    }
  }

  if (!out.length) out.push('gemini-3.1-pro-preview');
  return out;
}

function prioritizeModels(models) {
  const arr = models.slice();
  arr.sort((a, b) => scoreModel(b) - scoreModel(a));
  return arr;
}

function scoreModel(m) {
  const x = String(m || '').toLowerCase();
  let s = 0;
  if (x.indexOf('pro') !== -1) s += 100;
  if (x.indexOf('3.1') !== -1) s += 80;
  else if (x.indexOf('2.5') !== -1) s += 60;
  else if (x.indexOf('2.') !== -1) s += 40;
  if (x.indexOf('preview') !== -1) s += 10;
  if (x.indexOf('flash-lite') !== -1) s -= 20;
  else if (x.indexOf('flash') !== -1) s -= 10;
  return s;
}

function extractCandidateTextFromRaw(raw) {
  try {
    const root = JSON.parse(raw || '{}');
    const candidates = Array.isArray(root.candidates) ? root.candidates : [];
    for (let i = 0; i < candidates.length; i++) {
      const parts = (((candidates[i] || {}).content || {}).parts) || [];
      for (let p = 0; p < parts.length; p++) {
        const t = String((parts[p] && parts[p].text) || '').trim();
        if (t) return t;
      }
    }
    return '';
  } catch (e) {
    return '';
  }
}

function httpReq(method, url, body) {
  return new Promise((resolve) => {
    const u = new URL(url);
    const req = https.request({
      method: method,
      hostname: u.hostname,
      path: u.pathname + u.search,
      headers: { 'Content-Type': 'application/json' },
      timeout: 25000
    }, (res) => {
      let data = '';
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', () => resolve({ status: Number(res.statusCode || 0), body: data }));
    });

    req.on('timeout', () => {
      req.destroy(new Error('timeout'));
    });

    req.on('error', (err) => resolve({ status: 0, body: String(err && err.message ? err.message : err) }));

    if (body) req.write(JSON.stringify(body));
    req.end();
  });
}

async function httpReqWithRetry(method, url, body, maxRetries) {
  const tries = Number(maxRetries || LIVE_MAX_RETRIES);
  let res = null;
  for (let i = 1; i <= tries; i++) {
    res = await httpReq(method, url, body);
    if (res && Number(res.status || 0) === 200) return res;
    const code = Number((res && res.status) || 0);
    if (!LIVE_RETRYABLE_STATUS[code] || i >= tries) break;
    await sleepMs(LIVE_BASE_DELAY_MS * i + rand(40, 160));
  }
  return res || { status: 0, body: 'retry-exhausted' };
}

function sleepMs(ms) {
  return new Promise((resolve) => setTimeout(resolve, Math.max(0, Number(ms) || 0)));
}
function makeGasContext(initialProps) {
  const props = new Map(Object.entries(initialProps || {}).map(([k, v]) => [k, String(v)]));
  const cache = new Map();

  return {
    console,
    Math,
    Date,
    JSON,
    Logger: { log() {} },
    SpreadsheetApp: {
      BorderStyle: { SOLID: 'SOLID' },
      ProtectionType: { RANGE: 'RANGE', SHEET: 'SHEET' },
      newDataValidation: () => ({ requireValueInList() { return this; }, setAllowInvalid() { return this; }, build() { return {}; } }),
      newConditionalFormatRule: () => ({ whenFormulaSatisfied() { return this; }, setBackground() { return this; }, setRanges() { return this; }, build() { return {}; } }),
      getUi: () => ({
        alert() {},
        createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() { return this; } }),
        prompt: () => ({ getSelectedButton: () => 'CANCEL', getResponseText: () => '' }),
        ButtonSet: { OK_CANCEL: 'OK_CANCEL' },
        Button: { OK: 'OK' }
      }),
      getActive: () => ({ getSpreadsheetTimeZone: () => 'Europe/Madrid' }),
      getActiveSpreadsheet: () => ({ getSheets: () => [] })
    },
    ScriptApp: {
      getProjectTriggers: () => [],
      deleteTrigger() {},
      newTrigger: () => ({ forSpreadsheet() { return this; }, onOpen() { return this; }, timeBased() { return this; }, after() { return this; }, create() {} }),
      getScriptId: () => 'audit-script-id'
    },
    LockService: {
      getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }),
      getScriptLock: () => ({ tryLock: () => true, releaseLock() {} })
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (props.has(k) ? props.get(k) : ''),
        setProperty: (k, v) => props.set(k, String(v)),
        setProperties: (o) => Object.keys(o || {}).forEach((k) => props.set(k, String(o[k]))),
        deleteProperty: (k) => props.delete(k),
        deleteAllProperties: () => props.clear()
      })
    },
    CacheService: {
      getScriptCache: () => ({
        get: (k) => (cache.has(k) ? cache.get(k) : null),
        put: (k, v) => cache.set(k, String(v)),
        remove: (k) => cache.delete(k)
      }),
      getUserCache: () => ({ put() {}, get: () => '' })
    },
    Session: { getScriptTimeZone: () => 'Europe/Madrid', getEffectiveUser: () => ({ getEmail: () => 'audit@example.com' }) },
    HtmlService: { createHtmlOutput: () => ({ setWidth() { return this; }, setHeight() { return this; } }) },
    UrlFetchApp: { fetch: () => ({ getResponseCode: () => 500, getContentText: () => '{}' }) },
    DriveApp: {},
    MailApp: { sendEmail() {} },
    Utilities: {
      formatDate: (d, tz, fmt) => {
        const dd = p2(d.getDate()); const mm = p2(d.getMonth() + 1); const yyyy = d.getFullYear();
        const hh = p2(d.getHours()); const mi = p2(d.getMinutes()); const ss = p2(d.getSeconds());
        if (fmt === 'dd/MM/yyyy') return dd + '/' + mm + '/' + yyyy;
        if (fmt === 'HH:mm:ss') return hh + ':' + mi + ':' + ss;
        if (fmt === 'dd/MM/yyyy HH:mm:ss') return dd + '/' + mm + '/' + yyyy + ' ' + hh + ':' + mi + ':' + ss;
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
  const bins = ['git', 'C:/Program Files/Git/cmd/git.exe', 'C:/Program Files/Git/bin/git.exe'];
  for (let i = 0; i < bins.length; i++) {
    try {
      execFileSync(bins[i], ['--version'], { stdio: 'ignore' });
      return bins[i];
    } catch (e) {}
  }
  throw new Error('git no encontrado');
}

function read(p) { return fs.readFileSync(p, 'utf8'); }
function readMaybe(p) { try { return fs.readFileSync(p, 'utf8'); } catch (e) { return null; } }
function assert(cond, msg) { if (!cond) throw new Error(msg || 'assert'); }
function rand(min, max) { return Math.floor(Math.random() * (max - min)) + min; }
function p2(n) { return String(n).padStart(2, '0'); }
function truncate(v, n) { const s = String(v || ''); return s.length <= n ? s : s.substring(0, n); }
