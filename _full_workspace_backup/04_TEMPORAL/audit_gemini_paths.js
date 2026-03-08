const fs = require('fs');
const vm = require('vm');
const file = 'C:/Users/elrub/Desktop/CARPETA CODEX/01_PROYECTOS/festivales-github/CRM_FESTIVALES_ENGINE.gs';
const code = fs.readFileSync(file, 'utf8');

let fetchMode = 'throw';
const context = {
  console,
  Utilities: { sleep: () => {} },
  UrlFetchApp: {
    fetch: () => {
      if (fetchMode === 'throw') throw new Error('network down');
      if (fetchMode === '404') return { getResponseCode: () => 404, getContentText: () => '{}' };
      if (fetchMode === '429') return { getResponseCode: () => 429, getContentText: () => '{}' };
      if (fetchMode === '500') return { getResponseCode: () => 500, getContentText: () => '{}' };
      if (fetchMode === '200bad') return { getResponseCode: () => 200, getContentText: () => '{"candidates":[]}' };
      if (fetchMode === '200ok') {
        return {
          getResponseCode: () => 200,
          getContentText: () => JSON.stringify({
            candidates: [{ content: { parts: [{ text: '{"email":"a@b.com","telefono":"+34 612 345 678","nombreContacto":"Ana","observaciones":"ok"}' }] } }]
          })
        };
      }
      return { getResponseCode: () => 400, getContentText: () => '{}' };
    }
  },
  CacheService: { getUserCache: () => ({ put: () => {}, get: () => null }) },
  SpreadsheetApp: {},
  ScriptApp: {},
  HtmlService: {},
  MailApp: {}
};

vm.createContext(context);
vm.runInContext(code, context, { filename: file });

function run(mode, expectNull) {
  fetchMode = mode;
  const out = context.invocarGeminiConFallback_('hola', { type: 'OBJECT', properties: {}, required: [] });
  const isNull = out === null;
  if (expectNull && !isNull) {
    throw new Error('Esperaba null en modo ' + mode);
  }
  if (!expectNull && isNull) {
    throw new Error('Esperaba objeto en modo ' + mode);
  }
}

run('throw', true);
run('404', true);
run('429', true);
run('500', true);
run('200bad', true);
run('200ok', false);

console.log('OK: rutas de error Gemini toleradas sin throws fatales.');

