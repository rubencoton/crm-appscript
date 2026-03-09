// =============================================================================
// CRM AYUDAS Y SUBVENCIONES - VERSION ESTABLE V2
// Lista para pegar en Google Apps Script
// =============================================================================

const CRM_CONFIG = {
  SHEET_CONCURSOS: 'CONCURSOS',
  SHEET_NUEVOS: 'NUEVOS CONCURSOS',
  SHEET_BANDAS: 'BANDAS',
  SHEET_CORREO: 'CORREO',
  MAX_EXECUTION_MS: 4.8 * 60 * 1000,
  MAX_RETRIES: 3,
  MAX_PDF_SIZE_MB: 7,
  LOG_CACHE_KEY: 'LIVE_LOGS',
  LOG_CACHE_SECONDS: 3600,
  MAX_LOG_LINES: 250,
  DEFAULT_MODELS: [
    'gemini-3.1-pro-preview',
    'gemini-2.5-pro',
    'gemini-2.5-flash',
    'gemini-1.5-pro'
  ]
};

const CRM_COL = {
  NOMBRE: 1,
  ESTADO: 2,
  INSCRIPCION: 3,
  FECHA_LIMITE: 4,
  FECHA_DESARROLLO: 5,
  TIPO_PREMIO: 6,
  DETALLE_PREMIO: 7,
  DIRIGIDO_A: 8,
  MUNICIPIO: 9,
  PROVINCIA: 10,
  PAIS: 11,
  LINK1: 12,
  LINK2: 13,
  LINK3: 14,
  EMAIL: 15,
  TELEFONO: 16,
  NOTAS: 17
};

const CRM_ESTADO = {
  REVISAR: 'REVISAR',
  REVISADO_IA: 'REVISADO IA',
  REVISADO_HUMANO: 'REVISADO HUMANO',
  NUEVO_DESCUBRIMIENTO: 'NUEVO DESCUBRIMIENTO'
};

const CRM_ESTADO_LEGACY = {
  NUEVO_DESCUBRIMIENTOS: 'NUEVO DESCUBRIMIENTOS'
};

const CRM_INSCRIPCION = {
  ABIERTA: 'ABIERTA',
  CERRADA: 'CERRADA',
  SIN_PUBLICAR: 'SIN PUBLICAR'
};

const CRM_TIPO_PREMIO = ['ECONOMICO', 'SERVICIO', 'ACTUACION', 'RESIDENCIA', 'VARIOS'];
const CRM_TIPO_PREMIO_SET = { ECONOMICO: true, SERVICIO: true, ACTUACION: true, RESIDENCIA: true, VARIOS: true };
const CRM_NO_DATA = 'No publicado / No localizado';

const GEMINI_API_KEY_FIJA = 'AIzaSyC2AnnQuFgKOR_qGNl4jTrsoWF672bnK0M';
const CRM_PASSWORD_FIJA = '+rubencoton26';
const DESARROLLADOR_APP = 'RUBEN COTON';
const FIRMA_APP = 'DESARROLLADOR: RUBEN COTON';

// -----------------------------------------------------------------------------
// 1) MENU Y CONFIGURACION
// -----------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 CRM: Ayudas')
    .addItem('🚀 Escaner total (con consola)', 'lanzarModoTotal')
    .addItem('🔍 Auditar matriz (con consola)', 'lanzarModoActualizar')
    .addItem('Nuevos + Radar (con consola)', 'lanzarModoNuevas')
    .addSeparator()
    .addItem('Preparar hoja para decidir', 'prepararHojaDecisionInteligente')
    .addItem('Actualizar panel de escenarios', 'actualizarPanelDecision')
    .addItem('Bloquear celdas calculadas', 'aplicarBloqueoCeldasCalculadas')
    .addItem('Quitar bloqueo de calculadas', 'quitarBloqueoCeldasCalculadas')
    .addItem('Activar modo instantaneo (sin temporizador)', 'activarModoInstantaneo')
    .addItem('Activar modo asincrono (con temporizador)', 'activarModoAsincrono')
    .addSeparator()
    .addItem('Ver codigo (desplegable)', 'mostrarVisorCodigoProyecto_')
    .addItem('📟 Ver estado tecnico', 'mostrarPanelEstadoTecnico_')
    .addSeparator()
    .addItem('📨 Enviar boletin a BANDAS', 'verificarYEnviarCorreos')
    .addSeparator()
    .addItem('🧹 Purgar estado/logs', 'purgarSistema')
    .addToUi();
}

function configurarSistema() {
  SpreadsheetApp.getUi().alert('Configuracion deshabilitada. Este CRM usa API y password fijas en el codigo.\n' + FIRMA_APP);
}

function buildScriptEditorUrl_() {
  return 'https://script.google.com/home/projects/' + encodeURIComponent(ScriptApp.getScriptId()) + '/edit';
}

function mostrarVisorCodigoProyecto_() {
  const html = HtmlService.createHtmlOutput(`<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { margin:0; background:#070d09; color:#d7ffe8; font-family:'Courier New', monospace; }
    .wrap { padding:12px; }
    .card { border:1px solid #0d5c33; border-radius:10px; background:#05110a; padding:12px; box-shadow:0 0 20px rgba(0,255,133,0.15); }
    h3 { margin:0 0 8px 0; color:#00ff85; }
    .sub { margin:0 0 10px 0; color:#93d6ad; font-size:12px; }
    .row { display:flex; gap:8px; align-items:center; margin-bottom:8px; }
    input, select { width:100%; box-sizing:border-box; padding:8px; border:1px solid #0d5c33; border-radius:8px; background:#000; color:#00ff85; }
    button { border:none; border-radius:8px; padding:9px 12px; font-weight:700; cursor:pointer; font-family:'Courier New', monospace; }
    .ok { background:#00ff85; color:#01290f; }
    .aux { background:#00b4ff; color:#032133; }
    .soft { background:#334155; color:#fff; }
    #meta { color:#89c5a2; font-size:12px; margin-bottom:8px; }
    #code { width:100%; height:430px; box-sizing:border-box; background:#010603; color:#d7ffe8; border:1px solid #0d5c33; border-radius:8px; padding:10px; font-size:12px; line-height:1.4; white-space:pre; overflow:auto; }
    #msg { min-height:18px; color:#ffd280; font-size:12px; margin-top:8px; white-space:pre-wrap; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card" id="login">
      <h3>🧾 VISOR DE CODIGO | RUBEN COTON</h3>
      <p class="sub">Introduce password para cargar el codigo con desplegable por archivo.</p>
      <input id="pwd" type="password" placeholder="Password del CRM">
      <div class="row">
        <button class="ok" onclick="cargar()">🔓 Cargar codigo</button>
        <button class="soft" onclick="google.script.host.close()">Cerrar</button>
      </div>
      <div id="msg"></div>
    </div>

    <div class="card" id="viewer" style="display:none;">
      <h3>🧠 CODIGO DEL PROYECTO</h3>
      <div id="meta"></div>
      <div class="row">
        <select id="sel" onchange="render()"></select>
      </div>
      <div class="row">
        <button class="aux" onclick="copiar()">📋 Copiar archivo</button>
        <button class="soft" onclick="abrirEditor()">🌐 Abrir editor</button>
      </div>
      <pre id="code"></pre>
    </div>
  </div>
  <script>
    let payload = null;
    let files = [];

    function cargar() {
      const pass = (document.getElementById('pwd').value || '').trim();
      const msg = document.getElementById('msg');
      msg.textContent = 'Cargando codigo...';
      google.script.run
        .withSuccessHandler(function(data) {
          payload = data || {};
          files = Array.isArray(payload.files) ? payload.files : [];
          if (!files.length) throw new Error('No hay archivos para mostrar.');

          document.getElementById('login').style.display = 'none';
          document.getElementById('viewer').style.display = 'block';

          const sel = document.getElementById('sel');
          sel.innerHTML = '';
          files.forEach(function(f, idx) {
            const opt = document.createElement('option');
            opt.value = String(idx);
            opt.textContent = (f.name || 'SIN_NOMBRE') + ' [' + (f.type || 'DESCONOCIDO') + ']';
            sel.appendChild(opt);
          });

          document.getElementById('meta').textContent =
            'Script ID: ' + (payload.scriptId || '') +
            ' | Archivos: ' + files.length +
            ' | Cargado: ' + (payload.fetchedAt || '');

          render();
        })
        .withFailureHandler(function(e) {
          msg.textContent = '❌ ' + (e && e.message ? e.message : e);
        })
        .obtenerCodigoProyectoSeguro_(pass);
    }

    function render() {
      const idx = Number(document.getElementById('sel').value || '0');
      const f = files[idx] || files[0] || {};
      document.getElementById('code').textContent = String(f.source || '');
    }

    function copiar() {
      const txt = document.getElementById('code').textContent || '';
      navigator.clipboard.writeText(txt).then(function(){ alert('✅ Archivo copiado'); });
    }

    function abrirEditor() {
      const url = payload && payload.editUrl ? payload.editUrl : '';
      if (!url) {
        alert('No hay URL de editor disponible.');
        return;
      }
      window.open(url, '_blank');
    }
  </script>
</body>
</html>`).setWidth(980).setHeight(700);

  SpreadsheetApp.getUi().showModelessDialog(html, 'VISOR DE CODIGO | ' + DESARROLLADOR_APP);
}

function obtenerCodigoProyectoSeguro_(pass) {
  if (!validarPasswordServidor(pass)) {
    throw new Error('Password incorrecta.');
  }
  return obtenerCodigoProyecto_();
}

function obtenerCodigoProyecto_() {
  const scriptId = ScriptApp.getScriptId();
  const url = 'https://script.googleapis.com/v1/projects/' + encodeURIComponent(scriptId) + '/content';
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  });

  const code = resp.getResponseCode();
  if (code !== 200) {
    const body = resp.getContentText() || '';
    throw new Error('No se pudo cargar el codigo automaticamente (HTTP ' + code + '). Revisa permisos de Apps Script API. ' + body.substring(0, 220));
  }

  const parsed = JSON.parse(resp.getContentText() || '{}');
  const files = (parsed.files || []).map(function(f) {
    return {
      name: f.name || 'SIN_NOMBRE',
      type: f.type || 'DESCONOCIDO',
      source: String(f.source || '')
    };
  }).sort(function(a, b) {
    return a.name.localeCompare(b.name);
  });

  return {
    scriptId: scriptId,
    editUrl: buildScriptEditorUrl_(),
    fetchedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
    files: files
  };
}

function mostrarPanelEstadoTecnico_() {
  const html = HtmlService.createHtmlOutput(`<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { margin:0; background:#070d09; color:#d7ffe8; font-family:'Courier New', monospace; }
    .wrap { padding:12px; }
    .card { border:1px solid #0d5c33; border-radius:10px; background:#05110a; padding:12px; box-shadow:0 0 20px rgba(0,255,133,0.15); }
    h3 { margin:0 0 8px 0; color:#00ff85; }
    .row { display:flex; gap:8px; margin-bottom:8px; }
    button { border:none; border-radius:8px; padding:9px 12px; font-weight:700; cursor:pointer; font-family:'Courier New', monospace; }
    .ok { background:#00ff85; color:#01290f; }
    .aux { background:#00b4ff; color:#032133; }
    pre { margin:0; width:100%; height:520px; overflow:auto; background:#010603; color:#d7ffe8; border:1px solid #0d5c33; border-radius:8px; padding:10px; font-size:12px; line-height:1.35; white-space:pre-wrap; }
    #meta { color:#89c5a2; font-size:12px; margin-bottom:8px; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <h3>📟 ESTADO TECNICO | RUBEN COTON</h3>
      <div id="meta">Cargando...</div>
      <div class="row">
        <button class="ok" onclick="cargar()">🔄 Refrescar</button>
        <button class="aux" onclick="copiar()">📋 Copiar JSON</button>
      </div>
      <pre id="out"></pre>
    </div>
  </div>
  <script>
    function cargar() {
      google.script.run.withSuccessHandler(function(data){
        document.getElementById('meta').textContent = 'Script ID: ' + (data.scriptId || '') + ' | Hora: ' + (data.now || '');
        document.getElementById('out').textContent = JSON.stringify(data, null, 2);
      }).withFailureHandler(function(e){
        document.getElementById('meta').textContent = 'Error';
        document.getElementById('out').textContent = '❌ ' + (e && e.message ? e.message : e);
      }).obtenerEstadoTecnico_();
    }

    function copiar() {
      const txt = document.getElementById('out').textContent || '';
      navigator.clipboard.writeText(txt).then(function(){ alert('✅ Estado copiado'); });
    }

    cargar();
  </script>
</body>
</html>`).setWidth(920).setHeight(690);

  SpreadsheetApp.getUi().showModelessDialog(html, 'ESTADO TECNICO | ' + DESARROLLADOR_APP);
}

function obtenerEstadoTecnico_() {
  const props = PropertiesService.getScriptProperties();
  const keys = [
    'RUN_MODE', 'PHASE', 'CURRENT_ROW', 'TOTAL_ROWS',
    'FASE_ACTUAL', 'IS_RUNNING', 'IS_DONE', 'TIME_OUT',
    'STOP_REQUESTED', 'IDX_EXISTING', 'IDX_NEW', 'IDX_MODEL', CRM_MODE_PROP
  ];
  const snapshot = {};
  for (let i = 0; i < keys.length; i++) {
    snapshot[keys[i]] = props.getProperty(keys[i]) || '';
  }

  return {
    now: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
    developer: DESARROLLADOR_APP,
    scriptId: ScriptApp.getScriptId(),
    editUrl: buildScriptEditorUrl_(),
    estado: getEstadoProgreso(),
    props: snapshot
  };
}
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const row = e.range.getRow();
    const col = e.range.getColumn();

    if (typeof actualizarEdicionInstantaneaCRM_ === 'function' && sh.getName() === CRM_CONFIG.SHEET_CONCURSOS && row >= 2) {
      if (actualizarEdicionInstantaneaCRM_(sh, row, col, SpreadsheetApp.getActive())) {
        return;
      }
    }

    if (sh.getName() !== CRM_CONFIG.SHEET_CONCURSOS) return;
    if (row < 2) return;

    if (col === CRM_COL.ESTADO) {
      const estado = normalizeEstado_(e.range.getValue());
      if (e.range.getValue() !== estado) {
        e.range.setValue(estado);
      }

      let bg = '#EFEFEF';
      let fc = '#333333';
      if (estado === CRM_ESTADO.REVISADO_IA) { bg = '#D9EAD3'; fc = '#1E4620'; }
      if (estado === CRM_ESTADO.REVISADO_HUMANO) { bg = '#D0E0F5'; fc = '#113A67'; }
      if (estado === CRM_ESTADO.NUEVO_DESCUBRIMIENTO) { bg = '#FFE9B8'; fc = '#704D00'; }
      e.range.setBackground(bg).setFontColor(fc);
    }

    if (col === CRM_COL.INSCRIPCION || col === CRM_COL.ESTADO) {
      const ins = sanitizeValue_(sh.getRange(row, CRM_COL.INSCRIPCION).getValue()).toUpperCase();
      aplicarFormatoFila_(sh, row, ins, SpreadsheetApp.getActive().getSpreadsheetTimeZone());
    }

    if (col === CRM_COL.ESTADO) {
      const estadoFinal = normalizeEstado_(sh.getRange(row, CRM_COL.ESTADO).getValue());
      let bg2 = '#EFEFEF';
      let fc2 = '#333333';
      if (estadoFinal === CRM_ESTADO.REVISADO_IA) { bg2 = '#D9EAD3'; fc2 = '#1E4620'; }
      if (estadoFinal === CRM_ESTADO.REVISADO_HUMANO) { bg2 = '#D0E0F5'; fc2 = '#113A67'; }
      if (estadoFinal === CRM_ESTADO.NUEVO_DESCUBRIMIENTO) { bg2 = '#FFE9B8'; fc2 = '#704D00'; }
      sh.getRange(row, CRM_COL.ESTADO).setBackground(bg2).setFontColor(fc2);
    }
  } catch (err) {
    logCRM_('onEdit aviso: ' + err.message, 'warning');
  }
}
function lanzarModoTotal() {
  mostrarConsolaSegura_('TOTAL', '🚀 ESCANER TOTAL');
}

function lanzarModoActualizar() {
  mostrarConsolaSegura_('UPDATE', '🔍 AUDITORIA MATRIZ');
}

function lanzarModoNuevas() {
  mostrarConsolaSegura_('NEW', '🛰️ NUEVOS + RADAR');
}

function continuarEscaner() {
  mostrarConsolaSegura_('CONTINUE', '⏯️ CONTINUAR ESCANER');
}

function solicitarParada() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('STOP_REQUESTED', 'TRUE');
  logCRM_('🛑 Peticion de parada recibida.', 'warning');
  return 'OK';
}

function purgarSistema() {
  const props = PropertiesService.getScriptProperties();
  props.deleteAllProperties();
  try {
    CacheService.getScriptCache().remove(CRM_CONFIG.LOG_CACHE_KEY);
  } catch (err) {
    // no-op
  }
  SpreadsheetApp.getUi().alert('🧹 Sistema purgado: propiedades y logs eliminados.');
}


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
    <h3>CONFIGURACION SEGURA - RUBEN COTON</h3>
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
  SpreadsheetApp.getUi().showModalDialog(html, 'CONFIGURACION CRM | ' + DESARROLLADOR_APP);
}

function guardarConfiguracionSegura_(apiKey, password, modelsCsv) {
  throw new Error('La configuracion manual esta deshabilitada en esta version.');
}

function mostrarConsolaSegura_(mode, titleText) {
  const html = HtmlService.createHtmlOutput(`<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { margin:0; background:#020603; color:#d2ffe3; font-family:'Courier New', monospace; }
    .wrap::before { content:''; position:fixed; inset:0; pointer-events:none; background:linear-gradient(to bottom, rgba(0,255,133,0.05) 50%, rgba(0,0,0,0) 50%); background-size:100% 4px; opacity:0.35; }
    .wrap { padding:14px; height:100vh; box-sizing:border-box; display:flex; flex-direction:column; }
    .panel { border:1px solid #0d5c33; border-radius:10px; background:#05110a; padding:14px; box-shadow:0 0 20px rgba(0,255,133,0.15); }
    #login { display:block; }
    #console { display:none; height:100%; }
    h2 { margin:0 0 8px 0; font-size:20px; color:#00ff85; text-shadow:0 0 10px rgba(0,255,133,0.55); }
    .sub { margin:0 0 12px 0; color:#85caa2; font-size:12px; }
    input { width:100%; box-sizing:border-box; padding:10px; border-radius:8px; border:1px solid #0d5c33; background:#000; color:#00ff85; margin-bottom:10px; }
    .row { display:flex; gap:8px; align-items:center; margin-bottom:8px; }
    button { cursor:pointer; border:none; border-radius:8px; padding:10px 12px; font-weight:700; }
    .ok { background:#00ff85; color:#01290f; }
    .danger { background:#ef4444; color:#fff; }
    .copy { background:#00b4ff; color:#001d2b; }
    .continue { background:#f59e0b; color:#1f1300; display:none; }
    .small { font-size:12px; color:#fca5a5; min-height:18px; }
    .top { display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; }
    .badge { background:#00ff85; color:#01290f; font-weight:700; padding:4px 8px; border-radius:999px; font-size:12px; }
    #fase { color:#b8ffd5; margin-bottom:8px; font-size:12px; }
    #terminal { flex:1; background:#010603; border:1px solid #0d5c33; border-radius:8px; padding:10px; overflow:auto; font-size:12px; white-space:pre-wrap; line-height:1.45; }
    .line-time { color:#5aa07a; }
    .line-info { color:#d2ffe3; }
    .line-warning { color:#facc15; }
    .line-error { color:#f97316; }
    .line-fatal { color:#fecaca; background:#7f1d1d; padding:0 4px; }
    .line-success { color:#22c55e; }
    .line-scan { color:#5de7ff; }
    .line-system { color:#c4b5fd; }
    .line-title { color:#f8fafc; font-weight:700; text-decoration:underline; }
    .bar { display:flex; gap:8px; margin-top:10px; }
    .bar button { flex:1; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="panel" id="login">
      <h2>${sanitizeHtml_(titleText)} - RUBEN COTON</h2>
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
        <div>CONSOLA EN VIVO - RUBEN COTON</div>
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
        document.getElementById('fase').textContent = 'Inicializando motor...';

        const run = google.script.run
          .withSuccessHandler(function(){
            iniciarPolling();
          })
          .withFailureHandler(function(e){
            document.getElementById('fase').textContent = 'Error al iniciar';
            document.getElementById('terminal').innerText = '❌ ' + (e && e.message ? e.message : e);
          });

        if (mode === 'CONTINUE') {
          run.iniciarEscanerContinuacionDesdePanel_();
        } else {
          run.iniciarEscanerDesdePanel_(mode);
        }
      }).withFailureHandler(function(e){
        err.textContent = 'Error: ' + (e && e.message ? e.message : e);
      }).validarPasswordServidor(pass);
    }

    function iniciarPolling() {
      if (poll) clearInterval(poll);
      google.script.run.withSuccessHandler(renderEstado).getEstadoProgreso();
      poll = setInterval(function(){
        google.script.run
          .withSuccessHandler(renderEstado)
          .withFailureHandler(function(e){
            document.getElementById('fase').textContent = 'Error de comunicacion';
            document.getElementById('terminal').innerText = '❌ Error de polling: ' + (e && e.message ? e.message : e);
          })
          .getEstadoProgreso();
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
      google.script.run
        .withSuccessHandler(function(){
          iniciarPolling();
        })
        .withFailureHandler(function(e){
          document.getElementById('fase').textContent = 'Error al reanudar';
          document.getElementById('terminal').innerText = '❌ ' + (e && e.message ? e.message : e);
        })
        .iniciarEscanerContinuacionDesdePanel_();
    }

    function apagar() {
      google.script.run
        .withFailureHandler(function(e){
          document.getElementById('terminal').innerText = '❌ No se pudo solicitar apagado: ' + (e && e.message ? e.message : e);
        })
        .solicitarParada();
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
  SpreadsheetApp.getUi().showModelessDialog(html, 'CONSOLA CRM | ' + DESARROLLADOR_APP);
}

function validarPasswordServidor(pass) {
  return String(pass || '').trim() === CRM_PASSWORD_FIJA;
}

function iniciarEscanerDesdePanel_(mode) {
  validarConfiguracionMinima_();
  const props = PropertiesService.getScriptProperties();
  const modeFinal = sanitizeValue_(mode).toUpperCase() || 'TOTAL';

  props.setProperty('RUN_MODE', modeFinal);
  props.setProperty('TIME_OUT', 'FALSE');
  props.setProperty('IS_DONE', 'FALSE');
  props.setProperty('STOP_REQUESTED', 'FALSE');
  props.setProperty('FASE_ACTUAL', 'Inicializando motor en segundo plano...');

  encolarEjecucionAsincrona_();
  logCRM_('Solicitud de escaneo recibida (modo ' + modeFinal + ').', 'system');
  return getEstadoProgreso();
}

function iniciarEscanerContinuacionDesdePanel_() {
  validarConfiguracionMinima_();
  const props = PropertiesService.getScriptProperties();
  const mode = sanitizeValue_(props.getProperty('RUN_MODE'));
  if (!mode) {
    throw new Error('No hay una ejecucion previa para continuar.');
  }

  props.setProperty('TIME_OUT', 'FALSE');
  props.setProperty('IS_DONE', 'FALSE');
  props.setProperty('STOP_REQUESTED', 'FALSE');
  props.setProperty('FASE_ACTUAL', 'Reanudando motor en segundo plano...');

  encolarEjecucionAsincrona_();
  logCRM_('Solicitud de continuacion recibida.', 'system');
  return getEstadoProgreso();
}

function encolarEjecucionAsincrona_() {
  if (typeof usarModoInstantaneoCRM_ === 'function' && usarModoInstantaneoCRM_()) {
    ejecutorAsincrono_();
    return;
  }

  limpiarTriggersEjecucion_();
  ScriptApp.newTrigger('ejecutorAsincrono_').timeBased().after(1000).create();
}

function ejecutorAsincrono_() {
  const props = PropertiesService.getScriptProperties();
  const mode = sanitizeValue_(props.getProperty('RUN_MODE')) || 'TOTAL';
  try {
    ejecutorMaestro(mode);
  } finally {
    limpiarTriggersEjecucion_();
  }
}

function limpiarTriggersEjecucion_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'ejecutorAsincrono_') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
function validarConfiguracionMinima_() {
  if (!GEMINI_API_KEY_FIJA || !CRM_PASSWORD_FIJA) {
    throw new Error('Faltan API o password fijas en el codigo.');
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

// -----------------------------------------------------------------------------
// 2) ESTADO Y LOGS
// -----------------------------------------------------------------------------

function getEstadoProgreso() {
  const p = PropertiesService.getScriptProperties();
  let logs = [];
  try {
    const raw = CacheService.getScriptCache().get(CRM_CONFIG.LOG_CACHE_KEY);
    logs = raw ? JSON.parse(raw) : [];
  } catch (err) {
    logs = [];
  }

  return {
    actual: Number(p.getProperty('CURRENT_ROW') || '0'),
    total: Number(p.getProperty('TOTAL_ROWS') || '0'),
    fase: p.getProperty('FASE_ACTUAL') || 'Sin iniciar',
    stop: p.getProperty('STOP_REQUESTED') === 'TRUE',
    done: p.getProperty('IS_DONE') === 'TRUE',
    timeout: p.getProperty('TIME_OUT') === 'TRUE',
    running: p.getProperty('IS_RUNNING') === 'TRUE',
    mode: p.getProperty('RUN_MODE') || '',
    phase: p.getProperty('PHASE') || '',
    logs: logs
  };
}

function logCRM_(message, type) {
  const lvl = type || 'info';
  Logger.log('[' + lvl.toUpperCase() + '] ' + message);
  try {
    const cache = CacheService.getScriptCache();
    const raw = cache.get(CRM_CONFIG.LOG_CACHE_KEY);
    const list = raw ? JSON.parse(raw) : [];
    const now = new Date();
    const t = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    list.push({ t: t, c: lvl, m: message });
    while (list.length > CRM_CONFIG.MAX_LOG_LINES) {
      list.shift();
    }
    cache.put(CRM_CONFIG.LOG_CACHE_KEY, JSON.stringify(list), CRM_CONFIG.LOG_CACHE_SECONDS);
  } catch (err) {
    // no-op
  }
}

// -----------------------------------------------------------------------------
// 3) ORQUESTADOR
// -----------------------------------------------------------------------------

function ejecutarConPassword_(mode, allowResume) {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const storedPass = CRM_PASSWORD_FIJA;

  const prompt = ui.prompt('Seguridad', 'Introduce la password del CRM:', ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  if ((prompt.getResponseText() || '').trim() !== storedPass) {
    ui.alert('Password incorrecta.\n' + FIRMA_APP);
    return;
  }

  if (!allowResume) {
    props.setProperty('TIME_OUT', 'FALSE');
  }
  ejecutorMaestro(mode);

  const st = getEstadoProgreso();
  if (st.timeout) {
    ui.alert('Pausa por tiempo de ejecucion. Usa "Continuar escaner".\n' + FIRMA_APP);
  } else if (st.stop) {
    ui.alert('Proceso detenido.\n' + FIRMA_APP);
  } else if (st.done) {
    ui.alert('Proceso completado.\n' + FIRMA_APP);
  }
}

function ejecutorMaestro(modeReq) {
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) {
    const msg = 'Hay otro proceso en ejecucion. Espera unos segundos.';
    logCRM_(msg, 'warning');
    try {
      SpreadsheetApp.getUi().alert(msg + '\n' + FIRMA_APP);
    } catch (uiErr) {
      // Puede ejecutarse por trigger y no tener contexto de UI.
    }
    return;
  }

  try {
    if (props.getProperty('IS_RUNNING') === 'TRUE') {
      logCRM_('Proceso ya en ejecucion. Se ignora nueva peticion.', 'warning');
      return;
    }

    const resume = props.getProperty('TIME_OUT') === 'TRUE';
    const mode = resume
      ? (props.getProperty('RUN_MODE') || modeReq || 'TOTAL')
      : (modeReq || props.getProperty('RUN_MODE') || 'TOTAL');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetConcursos = ss.getSheetByName(CRM_CONFIG.SHEET_CONCURSOS);
    const sheetNuevos = ss.getSheetByName(CRM_CONFIG.SHEET_NUEVOS);
    if (!sheetConcursos) {
      throw new Error('No existe la pestana CONCURSOS.');
    }

    if (!resume) {
      resetRunState_(props, mode, sheetConcursos, sheetNuevos);
      try {
        CacheService.getScriptCache().remove(CRM_CONFIG.LOG_CACHE_KEY);
      } catch (err) {
        // no-op
      }
    } else {
      props.setProperty('TIME_OUT', 'FALSE');
      logCRM_('Reanudando escaner en fase: ' + (props.getProperty('PHASE') || 'N/A'), 'system');
    }

    props.setProperty('IS_RUNNING', 'TRUE');
    props.setProperty('IS_DONE', 'FALSE');
    props.setProperty('STOP_REQUESTED', 'FALSE');

    const startTime = Date.now();
    aplicarDiseno_(sheetConcursos, sheetNuevos);

    const phase = props.getProperty('PHASE') || getInitialPhase_(mode);
    props.setProperty('PHASE', phase);

    if ((mode === 'TOTAL' || mode === 'UPDATE') && phase === 'EXISTING') {
      const s1 = procesarMatriz_(ss, sheetConcursos, props, startTime);
      if (s1 !== 'DONE') return;
      props.setProperty('PHASE', mode === 'UPDATE' ? 'DONE' : 'NEW');
    }

    if ((mode === 'TOTAL' || mode === 'NEW') && props.getProperty('PHASE') === 'NEW') {
      const s2 = procesarNuevos_(ss, sheetConcursos, sheetNuevos, props, startTime);
      if (s2 !== 'DONE') return;
      props.setProperty('PHASE', 'RADAR');
    }

    if ((mode === 'TOTAL' || mode === 'NEW') && props.getProperty('PHASE') === 'RADAR') {
      const s3 = procesarRadar_(ss, sheetConcursos, props, startTime);
      if (s3 !== 'DONE') return;
      props.setProperty('PHASE', 'DONE');
    }

    if (typeof onEscaneoFinalizadoDecision_ === 'function') {
      onEscaneoFinalizadoDecision_(ss);
    }
    props.setProperty('IS_DONE', 'TRUE');
    props.setProperty('FASE_ACTUAL', 'Completado');
    logCRM_('Proceso finalizado correctamente.', 'success');
  } catch (err) {
    logCRM_('Error fatal: ' + (err && err.message ? err.message : err), 'fatal');
    throw err;
  } finally {
    props.setProperty('IS_RUNNING', 'FALSE');
    lock.releaseLock();
  }
}

function getInitialPhase_(mode) {
  if (mode === 'UPDATE') return 'EXISTING';
  if (mode === 'NEW') return 'NEW';
  return 'EXISTING';
}

function resetRunState_(props, mode, sheetConcursos, sheetNuevos) {
  const total = calcularTotalObjetivos_(sheetConcursos, sheetNuevos, mode);
  props.setProperties({
    RUN_MODE: mode,
    PHASE: getInitialPhase_(mode),
    IDX_EXISTING: '0',
    IDX_NEW: '0',
    CURRENT_ROW: '0',
    TOTAL_ROWS: String(total),
    FASE_ACTUAL: 'Preparando escaner',
    STOP_REQUESTED: 'FALSE',
    IS_DONE: 'FALSE',
    TIME_OUT: 'FALSE',
    IDX_MODEL: '0'
  });
  logCRM_('Inicio de escaneo en modo: ' + mode, 'title');
}

function calcularTotalObjetivos_(sheetConcursos, sheetNuevos, mode) {
  let total = 0;
  if ((mode === 'TOTAL' || mode === 'UPDATE') && sheetConcursos.getLastRow() > 1) {
    const data = sheetConcursos.getRange(2, CRM_COL.NOMBRE, sheetConcursos.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      const name = sanitizeValue_(data[i][0]);
      const state = normalizeEstado_(data[i][1]);
      if (name && state !== CRM_ESTADO.REVISADO_HUMANO) total++;
    }
  }
  if ((mode === 'TOTAL' || mode === 'NEW') && sheetNuevos && sheetNuevos.getLastRow() > 1) {
    const links = sheetNuevos.getRange(2, 1, sheetNuevos.getLastRow() - 1, 1).getValues();
    for (let j = 0; j < links.length; j++) {
      if (sanitizeValue_(links[j][0])) total++;
    }
  }
  if (mode === 'TOTAL' || mode === 'NEW') total += 5;
  return total;
}

function checkBudget_(props, startTime) {
  if (props.getProperty('STOP_REQUESTED') === 'TRUE') return 'STOP';
  if (Date.now() - startTime > CRM_CONFIG.MAX_EXECUTION_MS) return 'TIMEOUT';
  return 'OK';
}

function bumpProgress_(props) {
  const curr = Number(props.getProperty('CURRENT_ROW') || '0') + 1;
  props.setProperty('CURRENT_ROW', String(curr));
}

function setPaused_(props, phaseMessage) {
  props.setProperty('TIME_OUT', 'TRUE');
  props.setProperty('FASE_ACTUAL', phaseMessage);
  props.setProperty('IS_RUNNING', 'FALSE');
  logCRM_('Pausa por limite de tiempo. Usa "Continuar escaner".', 'warning');
}

// -----------------------------------------------------------------------------
// 4) FASE 1 - MATRIZ CONCURSOS
// -----------------------------------------------------------------------------

function procesarMatriz_(ss, sheetConcursos, props, startTime) {
  if (sheetConcursos.getLastRow() <= 1) return 'DONE';
  const totalRows = sheetConcursos.getLastRow() - 1;
  const data = sheetConcursos.getRange(2, 1, totalRows, 17).getValues();
  let idx = Number(props.getProperty('IDX_EXISTING') || '0');
  const tz = ss.getSpreadsheetTimeZone();

  props.setProperty('FASE_ACTUAL', 'Fase 1: Auditoria matriz');

  for (let i = idx; i < data.length; i++) {
    const budget = checkBudget_(props, startTime);
    if (budget === 'STOP') {
      props.setProperty('FASE_ACTUAL', 'Detenido por usuario');
      logCRM_('Proceso detenido por usuario.', 'warning');
      return 'STOP';
    }
    if (budget === 'TIMEOUT') {
      setPaused_(props, 'Pausa de seguridad en Fase 1');
      props.setProperty('IDX_EXISTING', String(i));
      return 'TIMEOUT';
    }

    const rowN = i + 2;
    const row = data[i];
    const nombre = sanitizeValue_(row[CRM_COL.NOMBRE - 1]);
    if (!nombre) {
      props.setProperty('IDX_EXISTING', String(i + 1));
      continue;
    }

    const estadoActual = normalizeEstado_(row[CRM_COL.ESTADO - 1]);
    if (estadoActual === CRM_ESTADO.REVISADO_HUMANO) {
      props.setProperty('IDX_EXISTING', String(i + 1));
      continue;
    }

    bumpProgress_(props);
    props.setProperty('FASE_ACTUAL', 'Fase 1: ' + nombre);
    logCRM_('Analizando matriz: ' + nombre, 'scan');

    let updated = row.slice();
    let inscripcionFinal = CRM_INSCRIPCION.SIN_PUBLICAR;

    try {
      const fuentesWeb = collectUrls_(row[CRM_COL.LINK1 - 1], row[CRM_COL.LINK2 - 1], row[CRM_COL.LINK3 - 1]);
      const urlPrincipal = fuentesWeb.length ? fuentesWeb[0] : '';
      const webObj = extraerWebProfundo_(fuentesWeb);
      const prompt = construirPromptFila_(row, fuentesWeb);
      const ai = llamarIA_(prompt, webObj, false);

      const fechaLimite = normalizarFechaLimite_(ai ? ai.fechaLimite : '', row[CRM_COL.FECHA_LIMITE - 1]);
      inscripcionFinal = estadoInscripcionDesdeFecha_(fechaLimite);
      const mesDesarrollo = normalizarMesDesarrollo_(mergeValue_(ai ? ai.fechasDesarrollo : '', row[CRM_COL.FECHA_DESARROLLO - 1]));
      const tipoPremio = normalizarTipoPremio_(mergeValue_(ai ? ai.tipoPremio : '', row[CRM_COL.TIPO_PREMIO - 1]));

      let link1 = normalizarUrl_(ai ? ai.link1 : '');
      if (inscripcionFinal === CRM_INSCRIPCION.ABIERTA && !isFechaEstimada_(fechaLimite) && !link1) {
        link1 = urlPrincipal || normalizarUrl_(row[CRM_COL.LINK1 - 1]) || 'Busqueda manual requerida';
      }

      updated[CRM_COL.ESTADO - 1] = CRM_ESTADO.REVISADO_IA;
      updated[CRM_COL.INSCRIPCION - 1] = inscripcionFinal;
      updated[CRM_COL.FECHA_LIMITE - 1] = fechaLimite;
      updated[CRM_COL.FECHA_DESARROLLO - 1] = mesDesarrollo;
      updated[CRM_COL.TIPO_PREMIO - 1] = tipoPremio;
      updated[CRM_COL.DETALLE_PREMIO - 1] = mergeValue_(ai ? ai.detallePremio : '', row[CRM_COL.DETALLE_PREMIO - 1]);
      updated[CRM_COL.DIRIGIDO_A - 1] = mergeValue_(ai ? ai.dirigidoA : '', row[CRM_COL.DIRIGIDO_A - 1]);
      updated[CRM_COL.MUNICIPIO - 1] = mergeValue_(ai ? ai.municipio : '', row[CRM_COL.MUNICIPIO - 1]);
      updated[CRM_COL.PROVINCIA - 1] = mergeValue_(ai ? ai.provincia : '', row[CRM_COL.PROVINCIA - 1]);
      updated[CRM_COL.PAIS - 1] = mergeValue_(ai ? ai.pais : '', row[CRM_COL.PAIS - 1]) || 'España';
      updated[CRM_COL.LINK1 - 1] = link1 || CRM_NO_DATA;
      updated[CRM_COL.LINK2 - 1] = normalizarUrl_(ai ? ai.link2 : '') || normalizarUrl_(row[CRM_COL.LINK2 - 1]) || CRM_NO_DATA;
      updated[CRM_COL.LINK3 - 1] = normalizarUrl_(ai ? ai.link3 : '') || CRM_NO_DATA;
      updated[CRM_COL.EMAIL - 1] = mergeValue_(ai ? ai.email : '', row[CRM_COL.EMAIL - 1]);
      updated[CRM_COL.TELEFONO - 1] = mergeValue_(ai ? ai.telefono : '', row[CRM_COL.TELEFONO - 1]);
      updated[CRM_COL.NOTAS - 1] = mergeValue_(ai ? ai.notas : '', row[CRM_COL.NOTAS - 1]);

      const emailAI = sanitizeValue_(ai ? ai.email : '');
      const emailPrev = sanitizeValue_(row[CRM_COL.EMAIL - 1]);
      if (!isValidEmail_(emailAI) && isValidEmail_(emailPrev)) {
        const notaActual = sanitizeValue_(updated[CRM_COL.NOTAS - 1]);
        if (notaActual.toLowerCase().indexOf('heredado de fuentes historicas') === -1) {
          updated[CRM_COL.NOTAS - 1] = (notaActual ? (notaActual + ' | ') : '') + 'Email de contacto heredado de fuentes historicas.';
        }
      }

      sheetConcursos.getRange(rowN, 1, 1, 17).setValues([updated]).clearNote();
      if (ai && sanitizeValue_(ai._razonamiento_logico)) {
        const note = [
          'ESTADO CALCULADO: ' + inscripcionFinal,
          'FECHA LIMITE: ' + fechaLimite,
          '',
          'REPORTE ANALISIS:',
          sanitizeValue_(ai._razonamiento_logico)
        ].join('\n');
        sheetConcursos.getRange(rowN, CRM_COL.FECHA_LIMITE).setNote(note);
      }
      logCRM_('Fila actualizada: ' + nombre + ' -> ' + inscripcionFinal, 'success');
    } catch (err) {
      updated[CRM_COL.ESTADO - 1] = CRM_ESTADO.REVISADO_IA;
      inscripcionFinal = estadoInscripcionDesdeFecha_(row[CRM_COL.FECHA_LIMITE - 1]);
      updated[CRM_COL.INSCRIPCION - 1] = inscripcionFinal;
      sheetConcursos.getRange(rowN, 1, 1, 17).setValues([updated]);
      logCRM_('Error en fila "' + nombre + '": ' + err.message, 'error');
    }

    aplicarFormatoFila_(sheetConcursos, rowN, inscripcionFinal, tz);
    props.setProperty('IDX_EXISTING', String(i + 1));
  }

  props.setProperty('IDX_EXISTING', '0');
  return 'DONE';
}

function construirPromptFila_(row, urls) {
  const fuentes = (Array.isArray(urls) && urls.length) ? urls.join(' | ') : 'No disponible';
  const historico = [
    'Nombre: ' + sanitizeValue_(row[CRM_COL.NOMBRE - 1]),
    'Fecha limite historica: ' + displayCell_(row[CRM_COL.FECHA_LIMITE - 1]),
    'Mes desarrollo historico: ' + displayCell_(row[CRM_COL.FECHA_DESARROLLO - 1]),
    'Tipo premio historico: ' + displayCell_(row[CRM_COL.TIPO_PREMIO - 1]),
    'Detalle premio historico: ' + displayCell_(row[CRM_COL.DETALLE_PREMIO - 1]),
    'Dirigido a historico: ' + displayCell_(row[CRM_COL.DIRIGIDO_A - 1]),
    'Municipio historico: ' + displayCell_(row[CRM_COL.MUNICIPIO - 1]),
    'Provincia historica: ' + displayCell_(row[CRM_COL.PROVINCIA - 1]),
    'Pais historico: ' + displayCell_(row[CRM_COL.PAIS - 1]),
    'Email historico: ' + displayCell_(row[CRM_COL.EMAIL - 1]),
    'Telefono historico: ' + displayCell_(row[CRM_COL.TELEFONO - 1]),
    'Links para contraste: ' + fuentes
  ].join('\n');

  return (
    'Analiza este concurso con criterio humano, logico y abierto. Devuelve JSON estricto.\n' +
    'PRIORIDADES DE DATOS:\n' +
    '1) Si existe base oficial del ano actual, usa ese dato como fuente principal.\n' +
    '2) Si no existe dato oficial actual, usa historico y estima con logica.\n' +
    '3) fechaLimite nunca vacia: fecha oficial DD/MM/YYYY o "ESTIMADO: DD/MM/YYYY".\n' +
    '4) fechasDesarrollo solo mes (Enero..Diciembre) o "ESTIMADO: Mes".\n' +
    '5) tipoPremio SOLO: ECONOMICO, SERVICIO, ACTUACION, RESIDENCIA o VARIOS.\n' +
    '6) Prioriza link1 como bases actuales; link2 historico; link3 complementario.\n' +
    '7) Si falta cualquier dato, usa "' + CRM_NO_DATA + '".\n' +
    '8) En _razonamiento_logico explica de donde sale cada dato y por que fue estimado.\n\n' +
    historico
  );
}

// -----------------------------------------------------------------------------
// 5) FASE 2 - NUEVOS CONCURSOS
// -----------------------------------------------------------------------------

function procesarNuevos_(ss, sheetConcursos, sheetNuevos, props, startTime) {
  if (!sheetNuevos || sheetNuevos.getLastRow() <= 1) return 'DONE';

  props.setProperty('FASE_ACTUAL', 'Fase 2: Analisis de nuevos links');
  let idx = Number(props.getProperty('IDX_NEW') || '0');
  const rows = sheetNuevos.getLastRow() - 1;
  const vals = sheetNuevos.getRange(2, 1, rows, 2).getValues();
  const existentes = getSetNombres_(sheetConcursos);

  for (let i = idx; i < vals.length; i++) {
    const budget = checkBudget_(props, startTime);
    if (budget === 'STOP') {
      props.setProperty('FASE_ACTUAL', 'Detenido por usuario');
      return 'STOP';
    }
    if (budget === 'TIMEOUT') {
      setPaused_(props, 'Pausa de seguridad en Fase 2');
      props.setProperty('IDX_NEW', String(i));
      return 'TIMEOUT';
    }

    const rowN = i + 2;
    const rawUrl = sanitizeValue_(vals[i][0]);
    if (!rawUrl) {
      props.setProperty('IDX_NEW', String(i + 1));
      continue;
    }

    bumpProgress_(props);
    props.setProperty('FASE_ACTUAL', 'Fase 2: link fila ' + rowN);

    if (!isValidHttpUrl_(rawUrl)) {
      sheetNuevos.getRange(rowN, 2).setValue('URL invalida');
      props.setProperty('IDX_NEW', String(i + 1));
      continue;
    }

    try {
      const webObj = extraerWebProfundo_([rawUrl]);
      const ai = llamarIA_(
        'Extrae TODOS los datos de este concurso nuevo. URL de origen: ' + rawUrl,
        webObj,
        false
      );
      if (!ai || !sanitizeValue_(ai.nombreConcurso)) {
        sheetNuevos.getRange(rowN, 2).setValue('Error de lectura');
        props.setProperty('IDX_NEW', String(i + 1));
        continue;
      }

      const nombre = sanitizeValue_(ai.nombreConcurso);
      if (existentes.has(nombre.toUpperCase())) {
        sheetNuevos.getRange(rowN, 2).setValue('Duplicado');
        props.setProperty('IDX_NEW', String(i + 1));
        continue;
      }

      const fechaLimite = normalizarFechaLimite_(ai.fechaLimite, '');
      const inscripcion = estadoInscripcionDesdeFecha_(fechaLimite);
      const tipoPremio = normalizarTipoPremio_(ai.tipoPremio);
      let link1 = normalizarUrl_(ai.link1);
      if (inscripcion === CRM_INSCRIPCION.ABIERTA && !isFechaEstimada_(fechaLimite) && !link1) {
        link1 = rawUrl;
      }

      const newRow = [
        nombre,
        CRM_ESTADO.NUEVO_DESCUBRIMIENTO,
        inscripcion,
        fechaLimite,
        normalizarMesDesarrollo_(ai.fechasDesarrollo),
        tipoPremio,
        mergeValue_(ai.detallePremio, ''),
        mergeValue_(ai.dirigidoA, ''),
        mergeValue_(ai.municipio, ''),
        mergeValue_(ai.provincia, ''),
        mergeValue_(ai.pais, '') || 'España',
        link1 || CRM_NO_DATA,
        normalizarUrl_(ai.link2) || CRM_NO_DATA,
        normalizarUrl_(ai.link3) || CRM_NO_DATA,
        mergeValue_(ai.email, ''),
        mergeValue_(ai.telefono, ''),
        mergeValue_(ai.notas, '')
      ];

      const destRow = sheetConcursos.getLastRow() + 1;
      sheetConcursos.getRange(destRow, 1, 1, 17).setValues([newRow]).clearNote();
      if (sanitizeValue_(ai._razonamiento_logico)) {
        sheetConcursos
          .getRange(destRow, CRM_COL.FECHA_LIMITE)
          .setNote('REPORTE ANALISIS:\n' + sanitizeValue_(ai._razonamiento_logico));
      }
      aplicarFormatoFila_(sheetConcursos, destRow, inscripcion, ss.getSpreadsheetTimeZone());

      sheetNuevos.getRange(rowN, 1, 1, 2).clearContent();
      existentes.add(nombre.toUpperCase());
      logCRM_('Nuevo concurso agregado: ' + nombre, 'success');
    } catch (err) {
      sheetNuevos.getRange(rowN, 2).setValue('Error: ' + (err.message || err));
      logCRM_('Error fase nuevos (fila ' + rowN + '): ' + err.message, 'error');
    }

    props.setProperty('IDX_NEW', String(i + 1));
  }

  props.setProperty('IDX_NEW', '0');
  return 'DONE';
}

function getSetNombres_(sheetConcursos) {
  const set = {};
  if (sheetConcursos.getLastRow() <= 1) return new Set();
  const names = sheetConcursos.getRange(2, CRM_COL.NOMBRE, sheetConcursos.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    const n = sanitizeValue_(names[i][0]);
    if (n) set[n.toUpperCase()] = true;
  }
  return new Set(Object.keys(set));
}

// -----------------------------------------------------------------------------
// 6) FASE 3 - RADAR
// -----------------------------------------------------------------------------

function procesarRadar_(ss, sheetConcursos, props, startTime) {
  props.setProperty('FASE_ACTUAL', 'Fase 3: Radar de oportunidades');

  const objetivo = 5;
  const maxIntentos = 5;
  const existentes = getSetNombres_(sheetConcursos);
  const rowsToInsert = [];
  const notes = [];
  const states = [];
  let intentos = 0;

  while (rowsToInsert.length < objetivo && intentos < maxIntentos) {
    const budget = checkBudget_(props, startTime);
    if (budget === 'STOP') return 'STOP';
    if (budget === 'TIMEOUT') {
      setPaused_(props, 'Pausa de seguridad en Fase 3');
      return 'TIMEOUT';
    }

    intentos++;
    const faltan = objetivo - rowsToInsert.length;
    const nombresMuestra = Array.from(existentes).slice(0, 120).join(', ');
    const prompt = [
      'Busca EXACTAMENTE ' + faltan + ' concursos o ayudas musicales vigentes para bandas de cualquier parte del mundo.',
      'No repitas estos nombres: ' + (nombresMuestra || 'Ninguno'),
      'Prioriza resultados con links reales y verificables.',
      'Si falta un dato, usa "' + CRM_NO_DATA + '".'
    ].join('\n');

    const arr = llamarIA_(prompt, { type: 'none' }, true);
    if (!arr || !Array.isArray(arr) || arr.length === 0) {
      logCRM_('Radar intento ' + intentos + ' sin resultados.', 'warning');
      continue;
    }

    for (let i = 0; i < arr.length && rowsToInsert.length < objetivo; i++) {
      const r = arr[i] || {};
      const nombre = sanitizeValue_(r.nombreConcurso);
      if (!nombre || existentes.has(nombre.toUpperCase())) continue;

      const fechaLimite = normalizarFechaLimite_(r.fechaLimite, '');
      const ins = estadoInscripcionDesdeFecha_(fechaLimite);
      const tipo = normalizarTipoPremio_(r.tipoPremio);

      rowsToInsert.push([
        nombre,
        CRM_ESTADO.NUEVO_DESCUBRIMIENTO,
        ins,
        fechaLimite,
        normalizarMesDesarrollo_(r.fechasDesarrollo),
        tipo,
        mergeValue_(r.detallePremio, ''),
        mergeValue_(r.dirigidoA, ''),
        mergeValue_(r.municipio, ''),
        mergeValue_(r.provincia, ''),
        mergeValue_(r.pais, '') || 'España',
        normalizarUrl_(r.link1) || CRM_NO_DATA,
        normalizarUrl_(r.link2) || CRM_NO_DATA,
        normalizarUrl_(r.link3) || CRM_NO_DATA,
        mergeValue_(r.email, ''),
        mergeValue_(r.telefono, ''),
        mergeValue_(r.notas, '')
      ]);
      notes.push(sanitizeValue_(r._razonamiento_logico));
      states.push(ins);
      existentes.add(nombre.toUpperCase());
      bumpProgress_(props);
    }
  }

  while (Number(props.getProperty('CURRENT_ROW') || '0') < Number(props.getProperty('TOTAL_ROWS') || '0') && rowsToInsert.length < objetivo) {
    // Mantiene progresion visual cuando el radar no llega a 5.
    bumpProgress_(props);
    rowsToInsert.push([
      'Radar pendiente de validacion manual [' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm') + ' #' + (rowsToInsert.length + 1) + ']',
      CRM_ESTADO.NUEVO_DESCUBRIMIENTO,
      CRM_INSCRIPCION.SIN_PUBLICAR,
      'SIN PUBLICAR',
      CRM_NO_DATA,
      'VARIOS',
      CRM_NO_DATA,
      CRM_NO_DATA,
      CRM_NO_DATA,
      CRM_NO_DATA,
      'España',
      CRM_NO_DATA,
      CRM_NO_DATA,
      CRM_NO_DATA,
      CRM_NO_DATA,
      CRM_NO_DATA,
      'Insertado automaticamente por radar al no alcanzar 5 resultados verificables.'
    ]);
    notes.push('Resultado de relleno para mantener 5 entradas en radar.');
    states.push(CRM_INSCRIPCION.SIN_PUBLICAR);
  }

  if (rowsToInsert.length > 0) {
    const startRow = sheetConcursos.getLastRow() + 1;
    sheetConcursos.getRange(startRow, 1, rowsToInsert.length, 17).setValues(rowsToInsert);
    for (let j = 0; j < rowsToInsert.length; j++) {
      if (notes[j]) {
        sheetConcursos.getRange(startRow + j, CRM_COL.FECHA_LIMITE).setNote('Radar:\n' + notes[j]);
      }
      aplicarFormatoFila_(sheetConcursos, startRow + j, states[j], ss.getSpreadsheetTimeZone());
    }
    logCRM_('Radar agrego ' + rowsToInsert.length + ' concursos.', 'success');
  } else {
    logCRM_('Radar sin resultados insertables.', 'warning');
  }

  return 'DONE';
}

// -----------------------------------------------------------------------------
// 7) LECTOR WEB/PDF + GEMINI
// -----------------------------------------------------------------------------

function extraerWeb_(url) {
  const u = sanitizeValue_(url);
  if (!isValidHttpUrl_(u)) return { type: 'none' };

  let finalUrl = u;
  if (finalUrl.indexOf('drive.google.com/file/d/') !== -1) {
    const match = finalUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      finalUrl = 'https://drive.google.com/uc?export=download&id=' + match[1];
    }
  }

  try {
    const resp = UrlFetchApp.fetch(finalUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    const code = resp.getResponseCode();
    if (code < 200 || code >= 300) return { type: 'none' };

    const headers = resp.getHeaders() || {};
    const ct = String(headers['Content-Type'] || headers['content-type'] || '').toLowerCase();
    if (finalUrl.toLowerCase().indexOf('.pdf') !== -1 || ct.indexOf('application/pdf') !== -1 || ct.indexOf('octet-stream') !== -1) {
      const blob = resp.getBlob();
      const bytes = blob.getBytes();
      if (bytes.length > CRM_CONFIG.MAX_PDF_SIZE_MB * 1024 * 1024) {
        return { type: 'none' };
      }
      return {
        type: 'pdf',
        mimeType: 'application/pdf',
        data: Utilities.base64Encode(bytes),
        url: finalUrl
      };
    }

    const html = resp.getContentText();
    const clean = html
      .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, ' ')
      .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, ' ')
      .replace(/<[^>]*>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .substring(0, 100000);

    return { type: 'text', content: clean, url: finalUrl };
  } catch (err) {
    return { type: 'none' };
  }
}

function llamarIA_(prompt, webObj, asArray) {
  const apiKey = getApiKey_();
  const models = getModels_();
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const hoy = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const year = now.getFullYear();
  let idxModel = Number(props.getProperty('IDX_MODEL') || '0');
  if (idxModel < 0 || idxModel >= models.length) idxModel = 0;

  const systemInstruction = [
    'Eres un equipo experto en management musical y convocatorias.',
    'Fecha de hoy: ' + hoy + ' | Ano: ' + year,
    'Debes devolver JSON valido y completo.',
    'Reglas:',
    '1) fechaLimite nunca vacia. Si no existe fecha oficial, estima como "ESTIMADO: DD/MM/YYYY".',
    '2) fechasDesarrollo solo mes (ej: "Abril" o "ESTIMADO: Abril").',
    '3) tipoPremio solo: ECONOMICO, SERVICIO, ACTUACION, RESIDENCIA, VARIOS.',
    '4) Rellena municipio, provincia, pais cuando exista evidencia.',
    '5) Prioriza link1 como bases actuales, link2 historico y link3 complementario.',
    '6) No menciones que eres IA.',
    '7) Cuando no exista un dato, usa "' + CRM_NO_DATA + '".'
  ].join('\n');

  const parts = [{ text: prompt }];
  if (webObj && webObj.type === 'text' && sanitizeValue_(webObj.content)) {
    parts.push({ text: 'Contenido web:\n' + webObj.content });
  } else if (webObj && webObj.type === 'pdf' && sanitizeValue_(webObj.data)) {
    parts.push({ text: 'Analiza este PDF para extraer datos:' });
    parts.push({ inlineData: { mimeType: webObj.mimeType, data: webObj.data } });
  } else {
    parts.push({ text: 'No hay contenido externo confiable. Usa historico y marca estimado cuando aplique.' });
  }

  const schemaObj = {
    type: 'OBJECT',
    properties: {
      nombreConcurso: { type: 'STRING' },
      _razonamiento_logico: { type: 'STRING' },
      inscripcion: { type: 'STRING' },
      fechaLimite: { type: 'STRING' },
      fechasDesarrollo: { type: 'STRING' },
      tipoPremio: { type: 'STRING' },
      detallePremio: { type: 'STRING' },
      dirigidoA: { type: 'STRING' },
      municipio: { type: 'STRING' },
      provincia: { type: 'STRING' },
      pais: { type: 'STRING' },
      link1: { type: 'STRING' },
      link2: { type: 'STRING' },
      link3: { type: 'STRING' },
      email: { type: 'STRING' },
      telefono: { type: 'STRING' },
      notas: { type: 'STRING' }
    },
    required: [
      'nombreConcurso',
      '_razonamiento_logico',
      'inscripcion',
      'fechaLimite',
      'fechasDesarrollo',
      'tipoPremio',
      'detallePremio',
      'dirigidoA',
      'municipio',
      'provincia',
      'pais',
      'link1',
      'link2',
      'link3',
      'email',
      'telefono',
      'notas'
    ]
  };

  const responseSchema = asArray ? { type: 'ARRAY', items: schemaObj } : schemaObj;
  const payload = {
    systemInstruction: { parts: [{ text: systemInstruction }] },
    contents: [{ role: 'user', parts: parts }],
    generationConfig: {
      responseMimeType: 'application/json',
      temperature: 0.1,
      maxOutputTokens: 8192,
      responseSchema: responseSchema
    }
  };

  for (let attempt = 1; attempt <= CRM_CONFIG.MAX_RETRIES; attempt++) {
    if (PropertiesService.getScriptProperties().getProperty('STOP_REQUESTED') === 'TRUE') return null;

    const model = models[idxModel];
    const endpoint = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + encodeURIComponent(apiKey);
    try {
      const resp = UrlFetchApp.fetch(endpoint, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      const code = resp.getResponseCode();
      const text = resp.getContentText();

      if (code === 200) {
        const parsed = JSON.parse(text);
        if (!parsed.candidates || !parsed.candidates.length) return null;
        const out = (((parsed.candidates[0] || {}).content || {}).parts || [])[0];
        if (!out || !out.text) return null;
        const json = extractJsonBlock_(out.text);
        if (!json) return null;
        const obj = JSON.parse(json);
        props.setProperty('IDX_MODEL', String(idxModel));
        return obj;
      }

      let msg = text;
      try {
        msg = (JSON.parse(text).error || {}).message || text;
      } catch (err) {
        // no-op
      }

      if (code === 429) {
        Utilities.sleep(1500);
        continue;
      }
      if ((code === 404 || code === 403 || code >= 500) && models.length > 1) {
        idxModel = (idxModel + 1) % models.length;
        props.setProperty('IDX_MODEL', String(idxModel));
        logCRM_('Cambio de modelo por error HTTP ' + code + '. Nuevo modelo: ' + models[idxModel], 'warning');
        continue;
      }
      if (code === 400 && msg.toLowerCase().indexOf('key') !== -1) {
        throw new Error('GEMINI_API_KEY invalida o no autorizada.');
      }
      logCRM_('Respuesta IA no valida (HTTP ' + code + '): ' + msg, 'error');
    } catch (err) {
      logCRM_('Error de red al llamar IA: ' + err.message, 'error');
      Utilities.sleep(1200);
    }
  }

  return null;
}

function extractJsonBlock_(text) {
  const m = text.match(/\{[\s\S]*\}|\[[\s\S]*\]/);
  return m ? m[0] : '';
}

function getApiKey_() {
  const key = String(GEMINI_API_KEY_FIJA || '').trim();
  if (!key) {
    throw new Error('GEMINI_API_KEY_FIJA no configurada.');
  }
  return key;
}

function getModels_() {
  const props = PropertiesService.getScriptProperties();
  const custom = sanitizeValue_(props.getProperty('GEMINI_MODELS_CSV'));
  if (!custom) return CRM_CONFIG.DEFAULT_MODELS.slice();
  const arr = custom.split(',').map(function (s) { return s.trim(); }).filter(function (s) { return !!s; });
  return arr.length ? arr : CRM_CONFIG.DEFAULT_MODELS.slice();
}

// -----------------------------------------------------------------------------
// 8) BOLETIN A BANDAS
// -----------------------------------------------------------------------------

function verificarYEnviarCorreos() {
  mostrarDialogoBoletinSeguro_();
}

function mostrarDialogoBoletinSeguro_() {
  const html = HtmlService.createHtmlOutput(`<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: Arial, sans-serif; background:#0f1115; color:#fff; margin:0; padding:16px; }
    .card { background:#171b22; border:1px solid #2f3642; border-radius:12px; padding:14px; }
    h3 { margin:0 0 8px 0; }
    p { margin:0 0 10px 0; color:#b8c0cc; font-size:13px; }
    input { width:100%; box-sizing:border-box; padding:10px; border-radius:8px; border:1px solid #3d4554; background:#0f131a; color:#fff; margin-bottom:10px; }
    .row { display:flex; gap:8px; }
    button { cursor:pointer; border:none; border-radius:8px; padding:10px 12px; font-weight:700; }
    .primary { background:#22c55e; color:#06240f; flex:1; }
    .secondary { background:#334155; color:#fff; }
    #status { font-size:12px; margin-top:8px; min-height:18px; color:#93c5fd; white-space:pre-wrap; }
  </style>
</head>
<body>
  <div class="card">
    <h3>ENVIAR BOLETIN - RUBEN COTON</h3>
    <p>Introduce password (oculta) para confirmar el envio.</p>
    <input id="pwd" type="password" placeholder="Password del CRM">
    <div class="row">
      <button class="secondary" onclick="google.script.host.close()">Cancelar</button>
      <button class="primary" onclick="enviar()">🚀 Enviar</button>
    </div>
    <div id="status"></div>
  </div>
  <script>
    function enviar() {
      if (!confirm('Se enviara el boletin a todas las bandas con email valido usando solo concursos ABIERTA. ¿Continuar?\\n\\nDESARROLLADOR: RUBEN COTON')) return;
      const pwd = document.getElementById('pwd').value || '';
      const status = document.getElementById('status');
      status.textContent = 'Enviando...';
      google.script.run.withSuccessHandler(function(msg){
        status.textContent = msg;
      }).withFailureHandler(function(err){
        status.textContent = 'Error: ' + (err && err.message ? err.message : err);
      }).enviarBoletinSeguroDesdePanel_(pwd);
    }
  </script>
</body>
</html>`).setWidth(460).setHeight(290);
  SpreadsheetApp.getUi().showModalDialog(html, 'SEGURIDAD DE ENVIO | ' + DESARROLLADOR_APP);
}

function enviarBoletinSeguroDesdePanel_(pass) {
  if (!validarPasswordServidor(pass)) {
    throw new Error('Password incorrecta.');
  }
  enviarBoletin_();
  return '✅ Envio ejecutado. Revisa el aviso final del sistema.';
}

function enviarBoletin_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const shCorreo = ss.getSheetByName(CRM_CONFIG.SHEET_CORREO);
  const shBandas = ss.getSheetByName(CRM_CONFIG.SHEET_BANDAS);
  const shConcursos = ss.getSheetByName(CRM_CONFIG.SHEET_CONCURSOS);
  if (!shCorreo || !shBandas || !shConcursos) {
    ui.alert('Faltan pestanas requeridas: CORREO, BANDAS o CONCURSOS.\n' + FIRMA_APP);
    return;
  }

  const subject = sanitizeValue_(shCorreo.getRange('B3').getValue()) || 'Nuevas oportunidades musicales';
  const baseMsg = sanitizeValue_(shCorreo.getRange('B6').getValue()) || 'Te compartimos los concursos abiertos seleccionados para tu banda.';
  const tz = ss.getSpreadsheetTimeZone();

  const concursos = shConcursos.getDataRange().getValues();
  const bloques = [];
  for (let i = 1; i < concursos.length; i++) {
    const row = concursos[i];
    const ins = sanitizeValue_(row[CRM_COL.INSCRIPCION - 1]).toUpperCase();
    if (ins !== CRM_INSCRIPCION.ABIERTA) continue;
    bloques.push(construirBloqueConcursoHtml_(row, tz));
  }

  if (!bloques.length) {
    ui.alert('No hay concursos en estado ABIERTA. Envio cancelado.\n' + FIRMA_APP);
    return;
  }

  const lastBandas = shBandas.getLastRow();
  if (lastBandas <= 1) {
    ui.alert('No hay bandas cargadas.\n' + FIRMA_APP);
    return;
  }

  const dataBandas = shBandas.getRange(2, 1, lastBandas - 1, 3).getValues();
  const updates = [];
  const stamp = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  let sent = 0;

  for (let i = 0; i < dataBandas.length; i++) {
    const name = sanitizeValue_(dataBandas[i][0]);
    const email = sanitizeValue_(dataBandas[i][1]);
    const prev = dataBandas[i][2];

    if (!name) {
      updates.push([prev]);
      continue;
    }
    if (!isValidEmail_(email)) {
      updates.push([prev]);
      continue;
    }

    const body = [
      '<div style="font-family:Arial,sans-serif;max-width:760px;margin:auto;color:#222;">',
      '<h2 style="margin-bottom:8px;">Hola ' + escapeHtml_(name) + '</h2>',
      '<p style="line-height:1.6;">' + escapeHtml_(baseMsg) + '</p>',
      '<p style="font-size:13px;background:#fff8e1;border-left:4px solid #f9a825;padding:10px 12px;">',
      '<strong>Regla del CRM:</strong> solo enviamos concursos con inscripcion ABIERTA (desde 2 meses antes de la fecha limite, inclusive, hasta la fecha limite).',
      '</p>',
      bloques.join('\n'),
      '<p style="margin-top:28px;font-size:12px;color:#666;">Mensaje generado desde tu CRM de ayudas y subvenciones.</p>',
      '<p style="margin-top:4px;font-size:11px;color:#666;">DESARROLLADOR: RUBEN COTON</p>',
      '</div>'
    ].join('\n');

    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: body
      });
      sent++;
      updates.push(['Enviado: ' + stamp]);
    } catch (err) {
      updates.push([prev]);
      logCRM_('No se pudo enviar a ' + email + ': ' + err.message, 'error');
    }
  }

  shBandas.getRange(2, 3, updates.length, 1).setValues(updates);
  ui.alert('Boletin enviado. Exitosos: ' + sent + '.\n' + FIRMA_APP);
}

function construirBloqueConcursoHtml_(row, tz) {
  const nombre = escapeHtml_(row[CRM_COL.NOMBRE - 1]);
  const fechaLim = formatCellDate_(row[CRM_COL.FECHA_LIMITE - 1], tz);
  const desarrollo = escapeHtml_(displayCell_(row[CRM_COL.FECHA_DESARROLLO - 1]) || '-');
  const tipo = escapeHtml_(displayCell_(row[CRM_COL.TIPO_PREMIO - 1]) || '-');
  const detalle = escapeHtml_(displayCell_(row[CRM_COL.DETALLE_PREMIO - 1]) || '-');
  const dirigido = escapeHtml_(displayCell_(row[CRM_COL.DIRIGIDO_A - 1]) || '-');
  const municipio = escapeHtml_(displayCell_(row[CRM_COL.MUNICIPIO - 1]) || '-');
  const provincia = escapeHtml_(displayCell_(row[CRM_COL.PROVINCIA - 1]) || '-');
  const pais = escapeHtml_(displayCell_(row[CRM_COL.PAIS - 1]) || '-');
  const email = escapeHtml_(displayCell_(row[CRM_COL.EMAIL - 1]) || '-');
  const tel = escapeHtml_(displayCell_(row[CRM_COL.TELEFONO - 1]) || '-');
  const notas = escapeHtml_(displayCell_(row[CRM_COL.NOTAS - 1]) || '-');

  const links = [];
  const l1 = normalizarUrl_(row[CRM_COL.LINK1 - 1]);
  const l2 = normalizarUrl_(row[CRM_COL.LINK2 - 1]);
  const l3 = normalizarUrl_(row[CRM_COL.LINK3 - 1]);
  if (l1) links.push('<a href="' + escapeHtml_(l1) + '">Bases actuales</a>');
  if (l2) links.push('<a href="' + escapeHtml_(l2) + '">Anteriores</a>');
  if (l3) links.push('<a href="' + escapeHtml_(l3) + '">Extra</a>');

  return [
    '<div style="border:1px solid #ddd;border-radius:8px;padding:14px;margin:14px 0;">',
    '<h3 style="margin:0 0 8px 0;color:#8b0000;">' + nombre + '</h3>',
    '<p style="margin:5px 0;"><strong>Fecha limite:</strong> ' + escapeHtml_(fechaLim) + '</p>',
    '<p style="margin:5px 0;"><strong>Desarrollo:</strong> ' + desarrollo + '</p>',
    '<p style="margin:5px 0;"><strong>Premio:</strong> ' + tipo + ' - ' + detalle + '</p>',
    '<p style="margin:5px 0;"><strong>Dirigido a:</strong> ' + dirigido + '</p>',
    '<p style="margin:5px 0;"><strong>Ubicacion:</strong> ' + municipio + ', ' + provincia + ' (' + pais + ')</p>',
    '<p style="margin:5px 0;"><strong>Enlaces:</strong> ' + (links.length ? links.join(' | ') : 'No publicado') + '</p>',
    '<p style="margin:5px 0;font-size:12px;color:#555;"><strong>Contacto:</strong> ' + email + ' | ' + tel + '</p>',
    '<p style="margin:5px 0;font-size:12px;color:#555;"><strong>Notas:</strong> ' + notas + '</p>',
    '</div>'
  ].join('\n');
}

// -----------------------------------------------------------------------------
// 9) DISENO Y VALIDACIONES
// -----------------------------------------------------------------------------

function aplicarDiseno_(sheetConcursos, sheetNuevos) {
  const maxCols = 17;
  const header = sheetConcursos.getRange(1, 1, 1, maxCols);
  header
    .setBackground('#6B0018')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheetConcursos.setFrozenRows(1);
  sheetConcursos.setRowHeight(1, 42);

  if (sheetConcursos.getFilter()) {
    sheetConcursos.getFilter().remove();
  }
  sheetConcursos.getRange(1, 1, Math.max(2, sheetConcursos.getLastRow()), maxCols).createFilter();

  if (sheetConcursos.getLastRow() > 1) {
    sheetConcursos.getRange(2, 1, sheetConcursos.getLastRow() - 1, 17)
      .setBackground('#F2F2F2')
      .setFontColor('#212121');
  }

  aplicarValidaciones_(sheetConcursos);

  if (sheetNuevos) {
    const colsN = sheetNuevos.getMaxColumns();
    if (colsN > 0) {
      sheetNuevos
        .getRange(1, 1, 1, colsN)
        .setBackground('#6B0018')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      sheetNuevos.setFrozenRows(1);
    }
  }
}

function aplicarValidaciones_(sheetConcursos) {
  const maxRows = Math.max(sheetConcursos.getLastRow() + 300, 1000);
  const estados = [
    CRM_ESTADO.REVISAR,
    CRM_ESTADO.REVISADO_IA,
    CRM_ESTADO.REVISADO_HUMANO,
    CRM_ESTADO.NUEVO_DESCUBRIMIENTO
  ];
  const ruleEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(estados, true)
    .setAllowInvalid(false)
    .build();
  sheetConcursos.getRange(2, CRM_COL.ESTADO, maxRows - 1, 1).setDataValidation(ruleEstado);

  const ruleIns = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      [CRM_INSCRIPCION.ABIERTA, CRM_INSCRIPCION.CERRADA, CRM_INSCRIPCION.SIN_PUBLICAR],
      true
    )
    .setAllowInvalid(false)
    .build();
  sheetConcursos.getRange(2, CRM_COL.INSCRIPCION, maxRows - 1, 1).setDataValidation(ruleIns);
}

function aplicarFormatoFila_(sheet, rowN, inscripcion, tz) {
  const dataCols = 17;
  const rowRange = sheet.getRange(rowN, 1, 1, dataCols);
  rowRange
    .setFontFamily('Roboto')
    .setFontSize(10)
    .setFontWeight('normal')
    .setFontColor('#333333')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

  const ins = sanitizeValue_(inscripcion).toUpperCase();
  let bg = '#FFFFFF';
  if (ins === CRM_INSCRIPCION.ABIERTA) bg = '#D9EAD3';
  else if (ins === CRM_INSCRIPCION.CERRADA) bg = '#F4CCCC';
  else if (ins === CRM_INSCRIPCION.SIN_PUBLICAR) bg = '#FFF2CC';
  rowRange.setBackground(bg);

  sheet.getRange(rowN, CRM_COL.NOMBRE).setFontWeight('bold').setFontSize(11).setFontColor('#000000');
  sheet.getRange(rowN, CRM_COL.ESTADO, 1, 4).setHorizontalAlignment('center');
  sheet.getRange(rowN, CRM_COL.TIPO_PREMIO).setHorizontalAlignment('center');

  const linkRange = sheet.getRange(rowN, CRM_COL.LINK1, 1, 3);
  linkRange.setFontSize(8).setVerticalAlignment('top');
  const links = linkRange.getValues()[0];
  for (let i = 0; i < 3; i++) {
    const val = normalizarUrl_(links[i]);
    const c = sheet.getRange(rowN, CRM_COL.LINK1 + i);
    if (val) {
      c.setFontColor('#1a73e8').setFontLine('underline');
    } else {
      c.setFontColor('#333333').setFontLine('none');
    }
  }
}

// -----------------------------------------------------------------------------
// 10) REGLAS DE NEGOCIO Y NORMALIZACION
// -----------------------------------------------------------------------------

function normalizeEstado_(value) {
  const v = sanitizeValue_(value).toUpperCase();
  if (!v) return CRM_ESTADO.REVISAR;
  if (v === CRM_ESTADO.REVISADO_HUMANO) return CRM_ESTADO.REVISADO_HUMANO;
  if (v === CRM_ESTADO.REVISADO_IA) return CRM_ESTADO.REVISADO_IA;
  if (v === CRM_ESTADO.NUEVO_DESCUBRIMIENTO || v === CRM_ESTADO_LEGACY.NUEVO_DESCUBRIMIENTOS) {
    return CRM_ESTADO.NUEVO_DESCUBRIMIENTO;
  }
  if (v === CRM_ESTADO.REVISAR) return CRM_ESTADO.REVISAR;
  return CRM_ESTADO.REVISAR;
}

function normalizarTipoPremio_(value) {
  const raw = sanitizeValue_(value)
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
  if (!raw) return 'VARIOS';
  if (raw.indexOf('ECONOM') !== -1) return 'ECONOMICO';
  if (raw.indexOf('SERVICIO') !== -1) return 'SERVICIO';
  if (raw.indexOf('ACTUACION') !== -1 || raw.indexOf('ACTUACI') !== -1) return 'ACTUACION';
  if (raw.indexOf('RESIDENCIA') !== -1) return 'RESIDENCIA';
  return CRM_TIPO_PREMIO_SET[raw] ? raw : 'VARIOS';
}

function normalizarUrl_(value) {
  const v = sanitizeValue_(value);
  if (isValidHttpUrl_(v)) return v;
  return '';
}

function firstUrl_(a, b, c) {
  const urls = collectUrls_(a, b, c);
  return urls.length ? urls[0] : '';
}

function collectUrls_(a, b, c) {
  const raw = [a, b, c];
  const out = [];
  for (let i = 0; i < raw.length; i++) {
    const u = normalizarUrl_(raw[i]);
    if (u && out.indexOf(u) === -1) out.push(u);
  }
  return out;
}

function extraerWebProfundo_(urlList) {
  const lista = Array.isArray(urlList) ? urlList : collectUrls_(urlList);
  if (!lista.length) return { type: 'none' };

  const bloques = [];
  const fuentes = [];
  for (let i = 0; i < lista.length && i < 3; i++) {
    const web = extraerWeb_(lista[i]);
    if (!web || web.type === 'none') continue;
    if (web.type === 'pdf') return web;
    if (web.type === 'text' && sanitizeValue_(web.content)) {
      bloques.push(web.content);
      fuentes.push(web.url || lista[i]);
    }
  }

  if (!bloques.length) return { type: 'none' };
  return {
    type: 'text',
    content: bloques.join('\n\n').substring(0, 120000),
    url: fuentes.join(' | ')
  };
}
function sanitizeValue_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function mergeValue_(newValue, oldValue, fallbackValue) {
  const n = sanitizeValue_(newValue);
  const o = sanitizeValue_(oldValue);
  const fallback = fallbackValue === undefined ? CRM_NO_DATA : sanitizeValue_(fallbackValue);

  if (n) {
    const l = n.toLowerCase();
    if (l === 'no publicado' || l === 'sin publicar' || l === 'n/a' || l === 'null' || l === CRM_NO_DATA.toLowerCase()) {
      return o || fallback;
    }
    return n;
  }

  if (o) return o;
  return fallback;
}
function displayCell_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return sanitizeValue_(value);
}

function normalizarFechaLimite_(newValue, oldValue) {
  const newStr = sanitizeValue_(newValue);
  const oldStr = displayCell_(oldValue);
  const parsedNew = parseFechaLimite_(newStr);
  if (parsedNew) {
    const formatted = formatDate_(parsedNew.date);
    return parsedNew.estimated ? 'ESTIMADO: ' + formatted : formatted;
  }

  const parsedOld = parseFechaLimite_(oldStr);
  if (parsedOld) {
    return parsedOld.estimated ? 'ESTIMADO: ' + formatDate_(parsedOld.date) : formatDate_(parsedOld.date);
  }

  return 'SIN PUBLICAR';
}

function parseFechaLimite_(value) {
  if (value instanceof Date) return { date: value, estimated: false };
  const raw = sanitizeValue_(value);
  if (!raw) return null;
  const estimated = raw.toUpperCase().indexOf('ESTIMADO') !== -1;
  const clean = raw.replace(/ESTIMADO\s*:?\s*/i, '').trim();
  const match = clean.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (!match) return null;
  const d = Number(match[1]);
  const m = Number(match[2]) - 1;
  const y = Number(match[3]);
  const dt = new Date(y, m, d, 12, 0, 0, 0);
  if (isNaN(dt.getTime())) return null;
  return { date: dt, estimated: estimated };
}

function isFechaEstimada_(value) {
  return sanitizeValue_(value).toUpperCase().indexOf('ESTIMADO') !== -1;
}

function estadoInscripcionDesdeFecha_(fechaValue) {
  const parsed = parseFechaLimite_(fechaValue);
  if (!parsed) return CRM_INSCRIPCION.SIN_PUBLICAR;

  const fLimite = new Date(parsed.date.getFullYear(), parsed.date.getMonth(), parsed.date.getDate(), 23, 59, 59, 999);
  const inicioVentana = new Date(fLimite);
  inicioVentana.setMonth(inicioVentana.getMonth() - 2);
  inicioVentana.setHours(0, 0, 0, 0);

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  if (hoy > fLimite) return CRM_INSCRIPCION.CERRADA;
  if (hoy >= inicioVentana && hoy <= fLimite) return CRM_INSCRIPCION.ABIERTA;
  return CRM_INSCRIPCION.SIN_PUBLICAR;
}

function normalizarMesDesarrollo_(value) {
  const raw = sanitizeValue_(value);
  if (!raw) return 'No publicado';
  const isEstimated = raw.toUpperCase().indexOf('ESTIMADO') !== -1;
  const clean = raw.replace(/ESTIMADO\s*:?\s*/i, '').trim();

  const meses = [
    { k: ['enero', 'january'], v: 'Enero' },
    { k: ['febrero', 'february'], v: 'Febrero' },
    { k: ['marzo', 'march'], v: 'Marzo' },
    { k: ['abril', 'april'], v: 'Abril' },
    { k: ['mayo', 'may'], v: 'Mayo' },
    { k: ['junio', 'june'], v: 'Junio' },
    { k: ['julio', 'july'], v: 'Julio' },
    { k: ['agosto', 'august'], v: 'Agosto' },
    { k: ['septiembre', 'setiembre', 'september'], v: 'Septiembre' },
    { k: ['octubre', 'october'], v: 'Octubre' },
    { k: ['noviembre', 'november'], v: 'Noviembre' },
    { k: ['diciembre', 'december'], v: 'Diciembre' }
  ];

  const lower = clean.toLowerCase();
  for (let i = 0; i < meses.length; i++) {
    for (let j = 0; j < meses[i].k.length; j++) {
      if (lower.indexOf(meses[i].k[j]) !== -1) {
        return isEstimated ? 'ESTIMADO: ' + meses[i].v : meses[i].v;
      }
    }
  }

  const dateParsed = parseFechaLimite_(clean);
  if (dateParsed) {
    const month = [
      'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
      'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
    ][dateParsed.date.getMonth()];
    return isEstimated ? 'ESTIMADO: ' + month : month;
  }

  return isEstimated ? ('ESTIMADO: ' + clean) : clean;
}

function isValidHttpUrl_(url) {
  const u = sanitizeValue_(url);
  return /^https?:\/\/\S+/i.test(u);
}

function isValidEmail_(email) {
  const e = sanitizeValue_(email);
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);
}

function formatDate_(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function formatCellDate_(value, tz) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, tz || Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return sanitizeValue_(value) || 'No publicado';
}

function escapeHtml_(value) {
  return sanitizeValue_(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}





