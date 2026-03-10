// =============================================================================
// CRM AYUDAS Y SUBVENCIONES - VERSION ESTABLE V2
// Lista para pegar en Google Apps Script
// =============================================================================

/*
 CONTEXTO GENERAL (LEER ANTES DE TOCAR ESTE ARCHIVO)
 - Este CRM esta pensado para gestionar ayudas/subvenciones para bandas y artistas.
 - El objetivo principal es tomar decisiones rapidas con datos incompletos, sin bloquear el flujo.
 - Regla clave de negocio: la FECHA LIMITE DE INSCRIPCION es el dato mas importante.
 - Si no existe fecha oficial, el sistema estima una fecha y la marca como "ESTIMADO".
 - Estado de inscripcion:
   * ABIERTA  -> hoy entre (fecha limite - 3 meses) y fecha limite.
   * CERRADA  -> fecha limite en el pasado.
   * SIN PUBLICAR -> no hay evidencia suficiente para fechar.
 - UX esperada: al editar datos en la hoja, el resultado debe verse al instante.
 - Este archivo contiene la orquestacion principal, la capa IA y normalizacion de datos.
*/
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
  AI_BACKOFF_BASE_MS: 1200,
  AI_BACKOFF_MAX_MS: 9000,
  AI_WEB_TEXT_LIMIT_CHARS: 100000,
  AI_MODEL_CACHE_KEY: 'CRM_GEM_MODELS_V2',
  AI_MODEL_CACHE_SECONDS: 6 * 60 * 60,
  DEFAULT_MODELS: [
    'gemini-3.1-pro-preview',
    'gemini-2.5-pro',
    'gemini-3-pro-preview',
    'gemini-pro-latest',
    'gemini-2.5-flash'
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

// Campos obligatorios para aceptar una respuesta de IA como valida.
const AI_REQUIRED_FIELDS = [
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
];

// IMPORTANTE: no dejar secretos reales en repositorio publico.
// Estas constantes son solo fallback local y deben quedarse vacias en git.
const GEMINI_API_KEY_FIJA = '';
const CRM_PASSWORD_FIJA = '';
const DESARROLLADOR_APP = 'RUBEN COTON';
const FIRMA_APP = 'DESARROLLADOR: RUBEN COTON';

// -----------------------------------------------------------------------------
// 1) MENU Y CONFIGURACION
// -----------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 CRM: Ayudas')
    .addItem('⚙️ Configurar API y password', 'configurarSistema')
    .addSeparator()
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
  // Se habilita configuracion segura para evitar credenciales hardcodeadas.
  mostrarPanelConfiguracion_();
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

  const parsed = safeJsonParse_(resp.getContentText() || '{}', {});
  const files = (Array.isArray(parsed.files) ? parsed.files : []).map(function(f) {
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
// onEdit es el motor de "respuesta instantanea" en la hoja.
// Se protege con lock para evitar pisados cuando hay varias ediciones simultaneas.
// Regla: solo procesa ediciones de una sola celda en CONCURSOS para minimizar riesgo de bucles.
function onEdit(e) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(250)) return;

  try {
    if (!e || !e.range) return;
    if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

    const ss = e.source || SpreadsheetApp.getActive();
    const sh = e.range.getSheet();
    const row = e.range.getRow();
    const col = e.range.getColumn();

    if (sh.getName() !== CRM_CONFIG.SHEET_CONCURSOS || row < 2) return;

    if (typeof actualizarEdicionInstantaneaCRM_ === 'function') {
      if (actualizarEdicionInstantaneaCRM_(sh, row, col, ss)) return;
    }

    if (col === CRM_COL.ESTADO) {
      const estado = normalizeEstado_(e.range.getValue());
      if (e.range.getValue() !== estado) {
        e.range.setValue(estado);
      }
      aplicarColorEstadoFila_(sh, row, estado);
    }

    if (col === CRM_COL.INSCRIPCION || col === CRM_COL.ESTADO || col === CRM_COL.FECHA_LIMITE) {
      const ins = sanitizeValue_(sh.getRange(row, CRM_COL.INSCRIPCION).getValue()).toUpperCase();
      aplicarFormatoFila_(sh, row, ins, ss.getSpreadsheetTimeZone());
    }
  } catch (err) {
    logCRM_('onEdit aviso: ' + err.message, 'warning');
  } finally {
    try {
      lock.releaseLock();
    } catch (lockErr) {
      // no-op
    }
  }
}

function aplicarColorEstadoFila_(sheet, row, estadoRaw) {
  const estado = normalizeEstado_(estadoRaw);
  let bg = '#EFEFEF';
  let fc = '#333333';
  if (estado === CRM_ESTADO.REVISADO_IA) { bg = '#D9EAD3'; fc = '#1E4620'; }
  if (estado === CRM_ESTADO.REVISADO_HUMANO) { bg = '#D0E0F5'; fc = '#113A67'; }
  if (estado === CRM_ESTADO.NUEVO_DESCUBRIMIENTO) { bg = '#FFE9B8'; fc = '#704D00'; }
  sheet.getRange(row, CRM_COL.ESTADO).setBackground(bg).setFontColor(fc);
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
  const props = PropertiesService.getScriptProperties();
  const key = sanitizeValue_(apiKey);
  const pass = sanitizeValue_(password);
  const models = sanitizeValue_(modelsCsv);

  if (!key) throw new Error('Debes indicar GEMINI_API_KEY.');
  if (!pass || pass.length < 6) throw new Error('La password debe tener al menos 6 caracteres.');

  props.setProperty('GEMINI_API_KEY', key);
  props.setProperty('CRM_PASSWORD', pass);

  if (models) props.setProperty('GEMINI_MODELS_CSV', models);
  else props.deleteProperty('GEMINI_MODELS_CSV');

  return 'Configuracion guardada correctamente.\\n' + FIRMA_APP;
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
  return sanitizeValue_(pass) === getPassword_();
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
  // Fuerza validacion temprana de credenciales para fallar con mensaje claro.
  getApiKey_();
  getPassword_();
}

function sanitizeHtml_(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// Utilidades defensivas: evitan que un JSON roto rompa la ejecucion.
function safeJsonParse_(raw, fallbackValue) {
  const txt = sanitizeValue_(raw);
  if (!txt) return fallbackValue;
  try {
    return JSON.parse(txt);
  } catch (err) {
    return fallbackValue;
  }
}

function truncateForLog_(value, maxLen) {
  const txt = sanitizeValue_(value);
  if (!txt) return '';
  const limit = Math.max(60, Number(maxLen) || 1200);
  if (txt.length <= limit) return txt;
  return txt.substring(0, limit) + '...';
}

// -----------------------------------------------------------------------------
// 2) ESTADO Y LOGS
// -----------------------------------------------------------------------------

function getEstadoProgreso() {
  const p = PropertiesService.getScriptProperties();
  let logs = [];
  try {
    const raw = CacheService.getScriptCache().get(CRM_CONFIG.LOG_CACHE_KEY);
    const parsed = raw ? safeJsonParse_(raw, []) : [];
    logs = Array.isArray(parsed) ? parsed : [];
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
    const parsed = raw ? safeJsonParse_(raw, []) : [];
    const list = Array.isArray(parsed) ? parsed : [];
    const now = new Date();
    const t = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    list.push({ t: t, c: lvl, m: truncateForLog_(message, 2000) });
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
  const storedPass = getPassword_();

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

// Construye el prompt de una fila usando historico + fuentes externas.
// Aqui se fuerza el criterio de negocio: priorizar fecha limite y explicacion razonada.
// El texto del prompt no es decorativo: define el comportamiento de la IA en produccion.
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
    '3) fechaLimite es el dato mas importante y nunca puede quedar vacia.\n' +
    '4) Si no hay fecha oficial clara, estimala como "ESTIMADO: DD/MM/YYYY".\n' +
    '5) Inscripcion ABIERTA cuando hoy este entre (fechaLimite - 3 meses) y fechaLimite, inclusive.\n' +
    '6) fechasDesarrollo solo mes (Enero..Diciembre) o "ESTIMADO: Mes".\n' +
    '7) tipoPremio SOLO: ECONOMICO, SERVICIO, ACTUACION, RESIDENCIA o VARIOS.\n' +
    '8) Prioriza link1 como bases actuales; link2 historico; link3 complementario.\n' +
    '9) Si falta cualquier dato, usa "' + CRM_NO_DATA + '".\n' +
    '10) En _razonamiento_logico explica de donde sale cada dato y por que fue estimado.\n' +
    '11) Indica contradicciones detectadas y nivel de confianza (ALTA/MEDIA/BAJA).\n\n' +
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

    const html = resp.getContentText() || '';
    if (!html) return { type: 'none' };

    const clean = html
      .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, ' ')
      .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, ' ')
      .replace(/<[^>]*>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .substring(0, CRM_CONFIG.AI_WEB_TEXT_LIMIT_CHARS);

    if (!clean) return { type: 'none' };
    return { type: 'text', content: clean, url: finalUrl };
  } catch (err) {
    return { type: 'none' };
  }
}

// Capa de llamada a IA con hardening:
// - Reintentos con backoff progresivo.
// - Rotacion de modelo cuando el endpoint responde con errores recuperables.
// - Parseo defensivo de JSON para evitar roturas por respuestas mal formadas.
// - Normalizacion obligatoria para devolver siempre una estructura util al CRM.
function llamarIA_(prompt, webObj, asArray) {
  const apiKey = getApiKey_();
  const models = getModels_(apiKey);
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const hoy = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const year = now.getFullYear();
  let idxModel = Number(props.getProperty('IDX_MODEL') || '0');
  if (idxModel < 0 || idxModel >= models.length) idxModel = 0;

  const systemInstruction = [
    'Eres un equipo experto en management musical y convocatorias.',
    'Fecha de hoy: ' + hoy + ' | Ano: ' + year,
    'Modelo prioritario: ' + (models[0] || 'gemini-3.1-pro-preview') + '.',
    'Debes devolver JSON valido y completo.',
    'Reglas:',
    '1) fechaLimite es el dato mas importante y nunca puede quedar vacia.',
    '2) Si no existe fecha oficial, estima de forma razonada como "ESTIMADO: DD/MM/YYYY".',
    '3) Inscripcion ABIERTA cuando hoy esta entre (fechaLimite - 3 meses) y fechaLimite, inclusive.',
    '4) fechasDesarrollo solo mes (ej: "Abril" o "ESTIMADO: Abril").',
    '5) tipoPremio solo: ECONOMICO, SERVICIO, ACTUACION, RESIDENCIA, VARIOS.',
    '6) Rellena municipio, provincia, pais cuando exista evidencia.',
    '7) Prioriza link1 como bases actuales, link2 historico y link3 complementario.',
    '8) Aplica razonamiento profundo: detecta contradicciones, explica supuestos y marca confianza ALTA/MEDIA/BAJA en _razonamiento_logico.',
    '9) Cuando no exista un dato, usa "' + CRM_NO_DATA + '".'
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
    required: AI_REQUIRED_FIELDS.slice()
  };

  const responseSchema = asArray ? { type: 'ARRAY', items: schemaObj } : schemaObj;
  const payload = {
    systemInstruction: { parts: [{ text: systemInstruction }] },
    contents: [{ role: 'user', parts: parts }],
    generationConfig: {
      responseMimeType: 'application/json',
      temperature: 0.05,
      maxOutputTokens: 8192,
      responseSchema: responseSchema
    }
  };

  for (let attempt = 1; attempt <= CRM_CONFIG.MAX_RETRIES; attempt++) {
    if (props.getProperty('STOP_REQUESTED') === 'TRUE') return null;

    const model = models[idxModel];
    const endpoint = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + encodeURIComponent(apiKey);

    try {
      const resp = UrlFetchApp.fetch(endpoint, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const code = Number(resp.getResponseCode() || 0);
      const textResp = String(resp.getContentText() || '');

      if (code === 200) {
        const parsed = safeJsonParse_(textResp, null);
        const candidateText = extractCandidateText_(parsed);
        const json = extractJsonBlock_(candidateText);
        const rawObj = safeJsonParse_(json, null);
        const normalized = normalizarRespuestaIA_(rawObj, !!asArray);

        if (normalized) {
          props.setProperty('IDX_MODEL', String(idxModel));
          return normalized;
        }

        logCRM_(
          'Respuesta IA sin estructura util. Intento ' + attempt + '/' + CRM_CONFIG.MAX_RETRIES +
            ' | Modelo: ' + model,
          'warning'
        );

        if (attempt < CRM_CONFIG.MAX_RETRIES) {
          Utilities.sleep(getBackoffMs_(attempt, code));
        }
        continue;
      }

      const msg = extractApiErrorMessage_(textResp);
      const msgLower = msg.toLowerCase();

      if (code === 400 && msgLower.indexOf('key') !== -1) {
        throw new Error('GEMINI_API_KEY invalida o no autorizada.');
      }

      if (shouldRotateModel_(code, msg) && models.length > 1) {
        idxModel = (idxModel + 1) % models.length;
        props.setProperty('IDX_MODEL', String(idxModel));
        logCRM_('Cambio de modelo por error HTTP ' + code + '. Nuevo modelo: ' + models[idxModel], 'warning');
      }

      if (shouldRetryIARequest_(code, msg) && attempt < CRM_CONFIG.MAX_RETRIES) {
        Utilities.sleep(getBackoffMs_(attempt, code));
        continue;
      }

      logCRM_('Respuesta IA no valida (HTTP ' + code + '): ' + truncateForLog_(msg, 350), 'error');
    } catch (err) {
      logCRM_(
        'Error de red al llamar IA: ' + truncateForLog_(err && err.message, 350) +
          ' | intento ' + attempt + '/' + CRM_CONFIG.MAX_RETRIES,
        'error'
      );
      if (attempt < CRM_CONFIG.MAX_RETRIES) {
        Utilities.sleep(getBackoffMs_(attempt, 0));
      }
    }
  }

  return null;
}

function extractJsonBlock_(text) {
  const raw = sanitizeValue_(text);
  if (!raw) return '';

  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenced && fenced[1]) {
    return fenced[1].trim();
  }

  const startObj = raw.indexOf('{');
  const startArr = raw.indexOf('[');
  let start = -1;
  if (startObj === -1) start = startArr;
  else if (startArr === -1) start = startObj;
  else start = Math.min(startObj, startArr);
  if (start < 0) return '';

  let depth = 0;
  let inString = false;
  let esc = false;
  for (let i = start; i < raw.length; i++) {
    const ch = raw[i];

    if (inString) {
      if (esc) {
        esc = false;
      } else if (ch === '\\') {
        esc = true;
      } else if (ch === '"') {
        inString = false;
      }
      continue;
    }

    if (ch === '"') {
      inString = true;
      continue;
    }

    if (ch === '{' || ch === '[') depth++;
    if (ch === '}' || ch === ']') depth--;

    if (depth === 0) {
      return raw.substring(start, i + 1).trim();
    }
  }

  return '';
}

function extractCandidateText_(parsed) {
  if (!parsed || !Array.isArray(parsed.candidates) || !parsed.candidates.length) return '';
  const cand = parsed.candidates[0] || {};
  const parts = ((cand.content || {}).parts || []);
  for (let i = 0; i < parts.length; i++) {
    const txt = sanitizeValue_(parts[i] && parts[i].text);
    if (txt) return txt;
  }
  return '';
}

function extractApiErrorMessage_(text) {
  const parsed = safeJsonParse_(text, null);
  if (parsed && parsed.error && parsed.error.message) {
    return String(parsed.error.message);
  }
  return sanitizeValue_(text) || 'Error sin detalle';
}

function shouldRetryIARequest_(code, msg) {
  if (code === 408 || code === 429 || code === 500 || code === 502 || code === 503 || code === 504) return true;
  const m = sanitizeValue_(msg).toLowerCase();
  return m.indexOf('timeout') !== -1 || m.indexOf('temporar') !== -1 || m.indexOf('overload') !== -1;
}

function shouldRotateModel_(code, msg) {
  if (code === 404 || code === 403) return true;
  if (code >= 500) return true;
  const m = sanitizeValue_(msg).toLowerCase();
  if (code === 400 && (m.indexOf('model') !== -1 || m.indexOf('not found') !== -1 || m.indexOf('unsupported') !== -1)) {
    return true;
  }
  return false;
}

function getBackoffMs_(attempt, code) {
  const base = Number(CRM_CONFIG.AI_BACKOFF_BASE_MS) || 1200;
  const max = Number(CRM_CONFIG.AI_BACKOFF_MAX_MS) || 9000;
  const att = Math.max(1, Number(attempt) || 1);
  const jitter = Math.floor(Math.random() * 250);
  const factor = (code === 429) ? 2.0 : 1.6;
  const ms = Math.floor(base * Math.pow(factor, att - 1)) + jitter;
  return Math.min(max, ms);
}

function normalizarRespuestaIA_(value, asArray) {
  if (asArray) {
    const src = Array.isArray(value) ? value : [value];
    const out = [];
    for (let i = 0; i < src.length; i++) {
      const obj = normalizarObjetoIA_(src[i]);
      if (obj) out.push(obj);
    }
    return out.length ? out : null;
  }

  const single = Array.isArray(value) ? value[0] : value;
  return normalizarObjetoIA_(single);
}

function normalizarObjetoIA_(raw) {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return null;

  const fechaLimite = normalizarFechaLimite_(raw.fechaLimite, '');
  const razonamiento = construirRazonamientoFuerte_(raw);

  const obj = {
    nombreConcurso: mergeValue_(raw.nombreConcurso, '', CRM_NO_DATA),
    _razonamiento_logico: razonamiento,
    inscripcion: normalizarInscripcionIA_(raw.inscripcion, fechaLimite),
    fechaLimite: fechaLimite,
    fechasDesarrollo: normalizarMesDesarrollo_(raw.fechasDesarrollo || fechaLimite),
    tipoPremio: normalizarTipoPremio_(raw.tipoPremio || 'VARIOS'),
    detallePremio: mergeValue_(raw.detallePremio, '', CRM_NO_DATA),
    dirigidoA: mergeValue_(raw.dirigidoA, '', CRM_NO_DATA),
    municipio: mergeValue_(raw.municipio, '', CRM_NO_DATA),
    provincia: mergeValue_(raw.provincia, '', CRM_NO_DATA),
    pais: mergeValue_(raw.pais, '', 'España'),
    link1: normalizarUrl_(raw.link1) || CRM_NO_DATA,
    link2: normalizarUrl_(raw.link2) || CRM_NO_DATA,
    link3: normalizarUrl_(raw.link3) || CRM_NO_DATA,
    email: sanitizeValue_(raw.email) || CRM_NO_DATA,
    telefono: sanitizeValue_(raw.telefono) || CRM_NO_DATA,
    notas: mergeValue_(raw.notas, '', CRM_NO_DATA)
  };

  for (let i = 0; i < AI_REQUIRED_FIELDS.length; i++) {
    const k = AI_REQUIRED_FIELDS[i];
    if (!sanitizeValue_(obj[k])) {
      obj[k] = (k === 'pais') ? 'España' : CRM_NO_DATA;
    }
  }

  return obj;
}

function normalizarInscripcionIA_(value, fechaLimiteNormalizada) {
  const v = sanitizeValue_(value).toUpperCase();
  if (v === CRM_INSCRIPCION.ABIERTA) return CRM_INSCRIPCION.ABIERTA;
  if (v === CRM_INSCRIPCION.CERRADA) return CRM_INSCRIPCION.CERRADA;
  if (v === CRM_INSCRIPCION.SIN_PUBLICAR) return CRM_INSCRIPCION.SIN_PUBLICAR;
  return estadoInscripcionDesdeFecha_(fechaLimiteNormalizada);
}

function construirRazonamientoFuerte_(raw) {
  const base = sanitizeValue_(raw && raw._razonamiento_logico);
  const urls = collectUrls_(raw && raw.link1, raw && raw.link2, raw && raw.link3);
  const fuentes = urls.length ? ('Fuentes: ' + urls.join(' | ')) : 'Fuentes: historico interno y estimacion.';

  let out = base || 'Sin evidencia oficial suficiente; se aplica estimacion conservadora para no bloquear decisiones.';
  if (out.length < 90) {
    out += ' | Se comparan bases actuales, historico y consistencia temporal.';
  }
  if (out.toUpperCase().indexOf('CONFIANZA') === -1) {
    out += ' | Confianza: MEDIA.';
  }
  out += ' | ' + fuentes;
  return truncateForLog_(out, 1800);
}

function getApiKey_() {
  const props = PropertiesService.getScriptProperties();
  const key = sanitizeValue_(props.getProperty('GEMINI_API_KEY')) || sanitizeValue_(GEMINI_API_KEY_FIJA);
  if (!key) {
    throw new Error('Falta GEMINI_API_KEY. Usa "Configurar API y password" en el menu.');
  }
  return key;
}

function getPassword_() {
  const props = PropertiesService.getScriptProperties();
  const pass = sanitizeValue_(props.getProperty('CRM_PASSWORD')) || sanitizeValue_(CRM_PASSWORD_FIJA);
  if (!pass) {
    throw new Error('Falta CRM_PASSWORD. Usa "Configurar API y password" en el menu.');
  }
  return pass;
}

function normalizeGeminiModelName_(name) {
  return sanitizeValue_(name).replace(/^models\//i, '');
}

function isGeminiTextModelCandidate_(name) {
  const model = normalizeGeminiModelName_(name).toLowerCase();
  if (!model || model.indexOf('gemini') !== 0) return false;
  if (/(image|audio|tts|embedding|aqa|veo|imagen|gemma|robotics|computer-use|deep-research)/i.test(model)) {
    return false;
  }
  return true;
}

function scoreGeminiModelPriority_(name) {
  const m = normalizeGeminiModelName_(name).toLowerCase();
  if (!isGeminiTextModelCandidate_(m)) return -999;
  let score = 0;
  if (m.indexOf('pro') !== -1) score += 100;
  if (m.indexOf('3.1') !== -1) score += 80;
  else if (m.indexOf('2.5') !== -1) score += 60;
  else if (m.indexOf('2.') !== -1) score += 40;
  if (m.indexOf('preview') !== -1) score += 10;
  if (m.indexOf('flash-lite') !== -1) score -= 20;
  else if (m.indexOf('flash') !== -1) score -= 10;
  return score;
}

function mergeModelLists_(discovered, customCsv, fallback) {
  const out = [];
  const seen = {};

  function pushModel_(name) {
    const model = normalizeGeminiModelName_(name);
    if (!model || seen[model]) return;
    if (!isGeminiTextModelCandidate_(model)) return;
    seen[model] = true;
    out.push(model);
  }

  for (let i = 0; i < discovered.length; i++) pushModel_(discovered[i]);

  const customArr = sanitizeValue_(customCsv)
    .split(',')
    .map(function (x) { return sanitizeValue_(x); })
    .filter(function (x) { return !!x; });
  for (let i = 0; i < customArr.length; i++) pushModel_(customArr[i]);

  for (let i = 0; i < fallback.length; i++) pushModel_(fallback[i]);

  if (!out.length) out.push('gemini-3.1-pro-preview');

  out.sort(function (a, b) {
    return scoreGeminiModelPriority_(b) - scoreGeminiModelPriority_(a);
  });

  return out;
}

function getModels_(apiKey) {
  const props = PropertiesService.getScriptProperties();
  const customCsv = sanitizeValue_(props.getProperty('GEMINI_MODELS_CSV'));
  const fallback = Array.isArray(CRM_CONFIG.DEFAULT_MODELS) && CRM_CONFIG.DEFAULT_MODELS.length
    ? CRM_CONFIG.DEFAULT_MODELS.slice()
    : ['gemini-3.1-pro-preview'];

  const cacheKey = sanitizeValue_(CRM_CONFIG.AI_MODEL_CACHE_KEY) || 'CRM_GEM_MODELS_V2';
  const cacheTtl = Math.max(300, Number(CRM_CONFIG.AI_MODEL_CACHE_SECONDS) || (6 * 60 * 60));

  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    if (cached) {
      const parsed = safeJsonParse_(cached, []);
      if (Array.isArray(parsed) && parsed.length) {
        return mergeModelLists_(parsed, customCsv, fallback);
      }
    }
  } catch (errCacheRead) {
    // no-op
  }

  let discovered = [];
  try {
    const key = sanitizeValue_(apiKey);
    if (key) {
      const url = 'https://generativelanguage.googleapis.com/v1beta/models?key=' + encodeURIComponent(key);
      const res = UrlFetchApp.fetch(url, {
        method: 'get',
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'CRM-AYUDAS/1.0' }
      });

      if (Number(res.getResponseCode() || 0) === 200) {
        const root = safeJsonParse_(res.getContentText() || '{}', {});
        const models = Array.isArray(root.models) ? root.models : [];
        discovered = models
          .map(function (m) { return normalizeGeminiModelName_(m && m.name); })
          .filter(function (m) { return isGeminiTextModelCandidate_(m); });

        discovered.sort(function (a, b) {
          return scoreGeminiModelPriority_(b) - scoreGeminiModelPriority_(a);
        });
      }
    }
  } catch (errFetchModels) {
    // no-op
  }

  const out = mergeModelLists_(discovered, customCsv, fallback);

  try {
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(out), cacheTtl);
  } catch (errCacheWrite) {
    // no-op
  }

  return out;
}

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
      '<strong>Regla del CRM:</strong> solo enviamos concursos con inscripcion ABIERTA (desde 3 meses antes de la fecha limite, inclusive, hasta la fecha limite).',
      '</p>',
      '<p style="font-size:12px;color:#555;margin:6px 0 14px 0;">Si la fecha exacta no esta publicada, el CRM muestra "ESTIMADO" para ayudarte a decidir sin frenar el flujo.</p>',
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
  aplicarFormatoCondicionalCalidad_(sheetConcursos);

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
  if (ins === CRM_INSCRIPCION.ABIERTA) bg = '#E3FCEF';
  else if (ins === CRM_INSCRIPCION.CERRADA) bg = '#FDE2E2';
  else if (ins === CRM_INSCRIPCION.SIN_PUBLICAR) bg = '#FFF4CC';
  rowRange.setBackground(bg);
  marcarCamposCriticosFila_(sheet, rowN, bg);

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

function aplicarFormatoCondicionalCalidad_(sheetConcursos) {
  const maxRows = Math.max(sheetConcursos.getLastRow(), 2);
  const fullRange = sheetConcursos.getRange(2, 1, maxRows - 1, 17);

  const ruleFecha = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",$D2="")')
    .setBackground('#FDE68A')
    .setRanges([fullRange])
    .build();

  const ruleContacto = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",$O2="",$P2="")')
    .setBackground('#FECACA')
    .setRanges([fullRange])
    .build();

  const ruleLinks = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",$L2="",$M2="",$N2="")')
    .setBackground('#FED7AA')
    .setRanges([fullRange])
    .build();

  const ruleUbicacion = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",$I2="",$J2="")')
    .setBackground('#FEE2E2')
    .setRanges([fullRange])
    .build();

  sheetConcursos.setConditionalFormatRules([ruleFecha, ruleContacto, ruleLinks, ruleUbicacion]);
}

function esDatoFaltanteCritico_(value) {
  const v = sanitizeValue_(value).toUpperCase();
  return !v || v === CRM_NO_DATA.toUpperCase() || v === 'SIN PUBLICAR' || v === 'NO PUBLICADO';
}

function marcarCamposCriticosFila_(sheet, rowN, fallbackBg) {
  const bgBase = fallbackBg || '#FFFFFF';
  const bgError = '#FEE2E2';
  const bgWarn = '#FEF3C7';

  const fechaCell = sheet.getRange(rowN, CRM_COL.FECHA_LIMITE);
  fechaCell.setBackground(bgBase);
  if (esDatoFaltanteCritico_(fechaCell.getValue())) {
    fechaCell.setBackground(bgError).setFontWeight('bold');
  } else {
    fechaCell.setFontWeight('bold');
  }

  const tipoCell = sheet.getRange(rowN, CRM_COL.TIPO_PREMIO);
  tipoCell.setBackground(bgBase);
  if (esDatoFaltanteCritico_(tipoCell.getValue())) tipoCell.setBackground(bgWarn);

  const muniCell = sheet.getRange(rowN, CRM_COL.MUNICIPIO);
  const provCell = sheet.getRange(rowN, CRM_COL.PROVINCIA);
  muniCell.setBackground(bgBase);
  provCell.setBackground(bgBase);
  if (esDatoFaltanteCritico_(muniCell.getValue()) && esDatoFaltanteCritico_(provCell.getValue())) {
    muniCell.setBackground(bgWarn);
    provCell.setBackground(bgWarn);
  }

  const emailCell = sheet.getRange(rowN, CRM_COL.EMAIL);
  const telCell = sheet.getRange(rowN, CRM_COL.TELEFONO);
  emailCell.setBackground(bgBase);
  telCell.setBackground(bgBase);
  if (esDatoFaltanteCritico_(emailCell.getValue()) && esDatoFaltanteCritico_(telCell.getValue())) {
    emailCell.setBackground(bgError);
    telCell.setBackground(bgError);
  }

  const linksRange = sheet.getRange(rowN, CRM_COL.LINK1, 1, 3);
  linksRange.setBackground(bgBase);
  const links = linksRange.getValues()[0];
  const hasLink = links.some(function (v) { return !!normalizarUrl_(v); });
  if (!hasLink) {
    linksRange.setBackground(bgWarn);
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

// Normaliza fecha limite con prioridad de negocio:
// 1) Fecha nueva valida.
// 2) Si no, fecha previa valida.
// 3) Si no hay ninguna, estimacion por defecto para no bloquear decisiones.
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

  return estimarFechaLimitePorDefecto_();
}

function parseFechaLimite_(value) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    return {
      date: new Date(value.getFullYear(), value.getMonth(), value.getDate(), 12, 0, 0, 0),
      estimated: false
    };
  }
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
  // Validacion estricta para impedir que JS "autocorrija" fechas invalidas
  // (ej: 32/01 -> 01/02), que alteraria estados y decisiones del CRM.
  if (dt.getFullYear() !== y || dt.getMonth() !== m || dt.getDate() !== d) {
    return null;
  }
  return { date: dt, estimated: estimated };
}

function isFechaEstimada_(value) {
  return sanitizeValue_(value).toUpperCase().indexOf('ESTIMADO') !== -1;
}

function estimarFechaLimitePorDefecto_() {
  const dt = new Date();
  dt.setDate(dt.getDate() + 90);
  return 'ESTIMADO: ' + formatDate_(dt);
}

// Traduce una fecha limite a estado operativo de inscripcion.
// Esta funcion materializa la regla central del CRM (ventana activa de 3 meses).
function estadoInscripcionDesdeFecha_(fechaValue) {
  const parsed = parseFechaLimite_(fechaValue);
  if (!parsed) return CRM_INSCRIPCION.SIN_PUBLICAR;

  const fLimite = new Date(parsed.date.getFullYear(), parsed.date.getMonth(), parsed.date.getDate(), 23, 59, 59, 999);
  const inicioVentana = restarMesesSeguro_(fLimite, 3);
  inicioVentana.setHours(0, 0, 0, 0);

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  if (hoy > fLimite) return CRM_INSCRIPCION.CERRADA;
  if (hoy >= inicioVentana && hoy <= fLimite) return CRM_INSCRIPCION.ABIERTA;
  return CRM_INSCRIPCION.SIN_PUBLICAR;
}

function restarMesesSeguro_(dateObj, months) {
  const base = new Date(dateObj);
  const day = base.getDate();
  const targetFirst = new Date(base.getFullYear(), base.getMonth() - Number(months || 0), 1, 12, 0, 0, 0);
  const lastDay = new Date(targetFirst.getFullYear(), targetFirst.getMonth() + 1, 0).getDate();
  const safeDay = Math.min(day, lastDay);
  return new Date(targetFirst.getFullYear(), targetFirst.getMonth(), safeDay, 12, 0, 0, 0);
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

function formatSpanishPhone_(raw) {
  const src = sanitizeValue_(raw);
  if (!src) return '';

  let digits = src.replace(/[^\d+]/g, '');
  if (digits.charAt(0) === '+') {
    digits = '+' + digits.substring(1).replace(/\D/g, '');
  } else {
    digits = digits.replace(/\D/g, '');
  }

  if (digits.indexOf('00') === 0) digits = '+' + digits.substring(2);

  if (digits.indexOf('+34') === 0) {
    digits = digits.substring(3);
  } else if (digits.indexOf('34') === 0 && digits.length === 11) {
    digits = digits.substring(2);
  } else if (digits.length === 9) {
    // numero nacional sin prefijo
  } else {
    return '';
  }

  if (!/^\d{9}$/.test(digits)) return '';
  return '+34 ' + digits.substring(0, 3) + ' ' + digits.substring(3, 6) + ' ' + digits.substring(6);
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







