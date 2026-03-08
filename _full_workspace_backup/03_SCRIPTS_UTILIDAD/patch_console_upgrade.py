import re
from pathlib import Path

path = Path(r'C:\Users\elrub\Desktop\CARPETA CODEX\crm_apps_script_v2.gs')
text = path.read_text(encoding='utf-8')

# 1) onOpen
text = re.sub(
    r"function onOpen\(\) \{[\s\S]*?\n\}\n\nfunction configurarSistema\(\)",
    """function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 CRM Ayudas')
    .addItem('🔐 1) Configurar API y password', 'configurarSistema')
    .addSeparator()
    .addItem('🚀 2) Escaner total (con consola)', 'lanzarModoTotal')
    .addItem('🔍 3) Auditar matriz (con consola)', 'lanzarModoActualizar')
    .addItem('🛰️ 4) Nuevos + Radar (con consola)', 'lanzarModoNuevas')
    .addItem('⏯️ 5) Continuar escaner', 'continuarEscaner')
    .addSeparator()
    .addItem('📨 6) Enviar boletin a BANDAS', 'verificarYEnviarCorreos')
    .addSeparator()
    .addItem('🛑 7) Apagado de emergencia', 'solicitarParada')
    .addItem('🧹 8) Purgar estado/logs', 'purgarSistema')
    .addToUi();
}

function configurarSistema()""",
    text,
    count=1
)

# 2) configure body
text = re.sub(
    r"function configurarSistema\(\) \{[\s\S]*?\n\}\n\nfunction lanzarModoTotal\(\)",
    """function configurarSistema() {
  mostrarPanelConfiguracion_();
}

function lanzarModoTotal()""",
    text,
    count=1
)

# 3) launchers and continue
text = text.replace("function lanzarModoTotal() {\n  ejecutarConPassword_('TOTAL');\n}",
                    "function lanzarModoTotal() {\n  mostrarConsolaSegura_('TOTAL', '🚀 ESCANER TOTAL');\n}")
text = text.replace("function lanzarModoActualizar() {\n  ejecutarConPassword_('UPDATE');\n}",
                    "function lanzarModoActualizar() {\n  mostrarConsolaSegura_('UPDATE', '🔍 AUDITORIA MATRIZ');\n}")
text = text.replace("function lanzarModoNuevas() {\n  ejecutarConPassword_('NEW');\n}",
                    "function lanzarModoNuevas() {\n  mostrarConsolaSegura_('NEW', '🛰️ NUEVOS + RADAR');\n}")

text = re.sub(
    r"function continuarEscaner\(\) \{[\s\S]*?\n\}\n",
    """function continuarEscaner() {
  mostrarConsolaSegura_('CONTINUE', '⏯️ CONTINUAR ESCANER');
}
""",
    text,
    count=1
)

# 4) messages
text = text.replace("logCRM_('Peticion de parada recibida.', 'warning');", "logCRM_('🛑 Peticion de parada recibida.', 'warning');")
text = text.replace("SpreadsheetApp.getUi().alert('Peticion de parada registrada.');", "SpreadsheetApp.getUi().alert('🛑 Peticion de parada registrada.');")
text = text.replace("SpreadsheetApp.getUi().alert('Sistema purgado: propiedades y logs eliminados.');", "SpreadsheetApp.getUi().alert('🧹 Sistema purgado: propiedades y logs eliminados.');")

# 5) inject new UI functions after purgarSistema
insert_block = r'''

function mostrarPanelConfiguracion_() {
  const props = PropertiesService.getScriptProperties();
  const html = HtmlService.createHtmlOutput(`<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: Arial, sans-serif; background:#0f1115; color:#fff; margin:0; padding:18px; }
    .card { background:#171b22; border:1px solid #2f3642; border-radius:12px; padding:16px; }
    h3 { margin:0 0 8px 0; }
    p { margin:0 0 14px 0; color:#b8c0cc; font-size:13px; }
    label { font-size:12px; color:#cdd5df; display:block; margin-bottom:6px; }
    input { width:100%; box-sizing:border-box; padding:10px; border-radius:8px; border:1px solid #3d4554; background:#0f131a; color:#fff; margin-bottom:12px; }
    .row { display:flex; gap:8px; }
    button { cursor:pointer; border:none; border-radius:8px; padding:10px 12px; font-weight:700; }
    .primary { background:#22c55e; color:#06240f; flex:1; }
    .secondary { background:#334155; color:#fff; }
    #status { font-size:12px; margin-top:10px; min-height:18px; color:#93c5fd; white-space:pre-wrap; }
    .hint { font-size:11px; color:#94a3b8; margin-top:4px; margin-bottom:10px; }
  </style>
</head>
<body>
  <div class="card">
    <h3>🔐 Configuracion Segura</h3>
    <p>Las claves se guardan en Script Properties y no quedan visibles en el codigo.</p>
    <label>Gemini API Key</label>
    <input id="apiKey" type="password" placeholder="AIza..." value="${sanitizeHtml_(props.getProperty('GEMINI_API_KEY') || '')}">
    <div class="hint">Tip: queda oculta mientras escribes.</div>
    <label>Password interna del CRM</label>
    <input id="pwd" type="password" placeholder="Tu password de acceso" value="${sanitizeHtml_(props.getProperty('CRM_PASSWORD') || '')}">
    <label>Modelos (opcional, CSV)</label>
    <input id="models" type="text" placeholder="gemini-2.5-pro, gemini-2.5-flash" value="${sanitizeHtml_(props.getProperty('GEMINI_MODELS_CSV') || '')}">
    <div class="row">
      <button class="secondary" onclick="google.script.host.close()">Cerrar</button>
      <button class="primary" onclick="guardar()">💾 Guardar</button>
    </div>
    <div id="status"></div>
  </div>
  <script>
    function guardar() {
      const apiKey = document.getElementById('apiKey').value || '';
      const pwd = document.getElementById('pwd').value || '';
      const models = document.getElementById('models').value || '';
      const status = document.getElementById('status');
      status.textContent = 'Guardando...';
      google.script.run.withSuccessHandler(function(msg){
        status.textContent = msg;
      }).withFailureHandler(function(err){
        status.textContent = 'Error: ' + (err && err.message ? err.message : err);
      }).guardarConfiguracionSegura_(apiKey, pwd, models);
    }
  </script>
</body>
</html>`).setWidth(520).setHeight(430);
  SpreadsheetApp.getUi().showModalDialog(html, '🔐 Configuracion CRM');
}

function guardarConfiguracionSegura_(apiKey, password, modelsCsv) {
  const api = String(apiKey || '').trim();
  const pwd = String(password || '').trim();
  const models = String(modelsCsv || '').trim();
  if (!api || !pwd) {
    throw new Error('Debes completar API key y password.');
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('GEMINI_API_KEY', api);
  props.setProperty('CRM_PASSWORD', pwd);
  if (models) {
    props.setProperty('GEMINI_MODELS_CSV', models);
  } else {
    props.deleteProperty('GEMINI_MODELS_CSV');
  }
  logCRM_('🔐 Configuracion guardada por usuario.', 'system');
  return '✅ Configuracion guardada correctamente.';
}

function mostrarConsolaSegura_(mode, titleText) {
  const html = HtmlService.createHtmlOutput(`<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { margin:0; background:#090d12; color:#e5e7eb; font-family:Consolas, monospace; }
    .wrap { padding:14px; height:100vh; box-sizing:border-box; display:flex; flex-direction:column; }
    .panel { border:1px solid #263041; border-radius:10px; background:#111827; padding:14px; }
    #login { display:block; }
    #console { display:none; height:100%; }
    h2 { margin:0 0 8px 0; font-size:20px; color:#22d3ee; }
    .sub { margin:0 0 12px 0; color:#94a3b8; font-size:12px; }
    input { width:100%; box-sizing:border-box; padding:10px; border-radius:8px; border:1px solid #334155; background:#0b1220; color:#fff; margin-bottom:10px; }
    .row { display:flex; gap:8px; align-items:center; margin-bottom:8px; }
    button { cursor:pointer; border:none; border-radius:8px; padding:10px 12px; font-weight:700; }
    .ok { background:#22c55e; color:#06240f; }
    .danger { background:#ef4444; color:#fff; }
    .copy { background:#38bdf8; color:#06233a; }
    .continue { background:#f59e0b; color:#1f1300; display:none; }
    .small { font-size:12px; color:#fca5a5; min-height:18px; }
    .top { display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; }
    .badge { background:#0ea5e9; color:#04243a; font-weight:700; padding:4px 8px; border-radius:999px; font-size:12px; }
    #fase { color:#93c5fd; margin-bottom:8px; font-size:12px; }
    #terminal { flex:1; background:#020617; border:1px solid #1e293b; border-radius:8px; padding:10px; overflow:auto; font-size:12px; white-space:pre-wrap; line-height:1.45; }
    .line-time { color:#64748b; }
    .line-info { color:#cbd5e1; }
    .line-warning { color:#facc15; }
    .line-error { color:#f97316; }
    .line-fatal { color:#fecaca; background:#7f1d1d; padding:0 4px; }
    .line-success { color:#22c55e; }
    .line-scan { color:#67e8f9; }
    .line-system { color:#c4b5fd; }
    .line-title { color:#f8fafc; font-weight:700; text-decoration:underline; }
    .bar { display:flex; gap:8px; margin-top:10px; }
    .bar button { flex:1; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="panel" id="login">
      <h2>🛡️ ${sanitizeHtml_(titleText)}</h2>
      <p class="sub">Introduce password para iniciar. La entrada va oculta.</p>
      <input type="password" id="pwd" placeholder="Password del CRM">
      <div class="row">
        <button class="ok" onclick="autenticar()">▶️ Iniciar</button>
        <button onclick="google.script.host.close()">Cerrar</button>
      </div>
      <div id="err" class="small"></div>
    </div>

    <div class="panel" id="console">
      <div class="top">
        <div>💻 Consola en vivo</div>
        <div id="pct" class="badge">0%</div>
      </div>
      <div id="fase">Preparando...</div>
      <div id="terminal"></div>
      <div class="bar">
        <button class="copy" onclick="copiar()">📋 Copiar logs</button>
        <button id="btnContinue" class="continue" onclick="continuar()">⏯️ Continuar</button>
        <button class="danger" onclick="apagar()">🛑 Apagado</button>
      </div>
    </div>
  </div>
  <script>
    const mode = ${JSON.stringify(mode)};
    let poll = null;
    let rawLogs = '';

    function autenticar() {
      const pass = document.getElementById('pwd').value || '';
      const err = document.getElementById('err');
      err.textContent = '';
      google.script.run.withSuccessHandler(function(ok){
        if (!ok) {
          err.textContent = '❌ Password incorrecta';
          return;
        }
        document.getElementById('login').style.display = 'none';
        document.getElementById('console').style.display = 'block';
        if (mode === 'CONTINUE') {
          google.script.run.iniciarEscanerContinuacionDesdePanel_();
        } else {
          google.script.run.iniciarEscanerDesdePanel_(mode);
        }
        iniciarPolling();
      }).withFailureHandler(function(e){
        err.textContent = 'Error: ' + (e && e.message ? e.message : e);
      }).validarPasswordServidor(pass);
    }

    function iniciarPolling() {
      if (poll) clearInterval(poll);
      poll = setInterval(function(){
        google.script.run.withSuccessHandler(renderEstado).getEstadoProgreso();
      }, 1000);
    }

    function renderEstado(st) {
      const total = Number(st.total || 0);
      const actual = Number(st.actual || 0);
      let pct = total > 0 ? Math.floor((actual / total) * 100) : 0;
      if (pct > 100) pct = 100;
      if (st.done) pct = 100;
      document.getElementById('pct').textContent = pct + '%';
      document.getElementById('fase').textContent = st.fase || '...';

      const btnContinue = document.getElementById('btnContinue');
      if (st.timeout) {
        btnContinue.style.display = 'block';
      } else {
        btnContinue.style.display = 'none';
      }

      if (st.done || st.stop) {
        if (poll) clearInterval(poll);
      }

      const logs = Array.isArray(st.logs) ? st.logs : [];
      rawLogs = logs.map(function(l){ return '[' + l.t + '] [' + String(l.c || 'info').toUpperCase() + '] ' + l.m; }).join('\\n');
      const html = logs.map(function(l){
        const kind = String(l.c || 'info').toLowerCase();
        const cls = 'line-' + (kind === 'info' ? 'info' : kind);
        return '<span class="line-time">[' + l.t + ']</span> <span class="' + cls + '">' + escapeHtml(String(l.m || '')) + '</span>';
      }).join('<br>');
      const term = document.getElementById('terminal');
      if (term.innerHTML !== html) {
        term.innerHTML = html;
        term.scrollTop = term.scrollHeight;
      }
    }

    function continuar() {
      document.getElementById('fase').textContent = '⏳ Reanudando...';
      google.script.run.iniciarEscanerContinuacionDesdePanel_();
      iniciarPolling();
    }

    function apagar() {
      google.script.run.solicitarParada();
    }

    function copiar() {
      navigator.clipboard.writeText(rawLogs || '').then(function(){
        alert('✅ Logs copiados');
      });
    }

    function escapeHtml(t) {
      return t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\"/g, '&quot;');
    }
  </script>
</body>
</html>`).setWidth(760).setHeight(650);
  SpreadsheetApp.getUi().showModelessDialog(html, '💻 Consola CRM');
}

function validarPasswordServidor(pass) {
  const storedPass = String(PropertiesService.getScriptProperties().getProperty('CRM_PASSWORD') || '').trim();
  if (!storedPass) {
    throw new Error('Primero configura API y password en el menu.');
  }
  return String(pass || '').trim() === storedPass;
}

function iniciarEscanerDesdePanel_(mode) {
  validarConfiguracionMinima_();
  PropertiesService.getScriptProperties().setProperty('TIME_OUT', 'FALSE');
  ejecutorMaestro(mode);
  return getEstadoProgreso();
}

function iniciarEscanerContinuacionDesdePanel_() {
  validarConfiguracionMinima_();
  const mode = String(PropertiesService.getScriptProperties().getProperty('RUN_MODE') || '').trim();
  if (!mode) {
    throw new Error('No hay una ejecucion previa para continuar.');
  }
  ejecutorMaestro(mode);
  return getEstadoProgreso();
}

function validarConfiguracionMinima_() {
  const props = PropertiesService.getScriptProperties();
  const api = String(props.getProperty('GEMINI_API_KEY') || '').trim();
  const pwd = String(props.getProperty('CRM_PASSWORD') || '').trim();
  if (!api || !pwd) {
    throw new Error('Falta configuracion: abre "🔐 Configurar API y password".');
  }
}

function sanitizeHtml_(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
'''

text = text.replace(
"""function purgarSistema() {
  const props = PropertiesService.getScriptProperties();
  props.deleteAllProperties();
  try {
    CacheService.getScriptCache().remove(CRM_CONFIG.LOG_CACHE_KEY);
  } catch (err) {
    // no-op
  }
  SpreadsheetApp.getUi().alert('🧹 Sistema purgado: propiedades y logs eliminados.');
}
""",
"""function purgarSistema() {
  const props = PropertiesService.getScriptProperties();
  props.deleteAllProperties();
  try {
    CacheService.getScriptCache().remove(CRM_CONFIG.LOG_CACHE_KEY);
  } catch (err) {
    // no-op
  }
  SpreadsheetApp.getUi().alert('🧹 Sistema purgado: propiedades y logs eliminados.');
}
""" + insert_block,
1
)

path.write_text(text, encoding='utf-8')
print('patched')
