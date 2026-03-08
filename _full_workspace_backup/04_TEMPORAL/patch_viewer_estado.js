const fs = require('fs');
const path = 'C:/Users/elrub/Desktop/CARPETA CODEX/02_APPS_SCRIPT_FUENTES/crm_ayudas_prod/Code.gs';
let s = fs.readFileSync(path, 'utf8');

function replaceOnce(oldText, newText, label) {
  const i = s.indexOf(oldText);
  if (i < 0) throw new Error('No encontrado bloque: ' + label);
  s = s.slice(0, i) + newText + s.slice(i + oldText.length);
}

function insertAfter(anchor, insertText, label) {
  const i = s.indexOf(anchor);
  if (i < 0) throw new Error('No encontrado ancla: ' + label);
  const pos = i + anchor.length;
  s = s.slice(0, pos) + insertText + s.slice(pos);
}

// 1) Menu mejorado con visor de codigo + estado tecnico
replaceOnce(
`function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 CRM: Ayudas')
    .addItem('🚀 Escaner total (con consola)', 'lanzarModoTotal')
    .addItem('🔍 Auditar matriz (con consola)', 'lanzarModoActualizar')
    .addItem('🛰️ Nuevos + Radar (con consola)', 'lanzarModoNuevas')
    .addSeparator()
    .addItem('📨 Enviar boletin a BANDAS', 'verificarYEnviarCorreos')
    .addSeparator()
    .addItem('🧹 Purgar estado/logs', 'purgarSistema')
    .addToUi();
}`,
`function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 CRM: Ayudas')
    .addItem('🚀 Escaner total (con consola)', 'lanzarModoTotal')
    .addItem('🔍 Auditar matriz (con consola)', 'lanzarModoActualizar')
    .addItem('🛰️ Nuevos + Radar (con consola)', 'lanzarModoNuevas')
    .addSeparator()
    .addItem('🧾 Ver codigo (desplegable)', 'mostrarVisorCodigoProyecto_')
    .addItem('📟 Ver estado tecnico', 'mostrarPanelEstadoTecnico_')
    .addSeparator()
    .addItem('📨 Enviar boletin a BANDAS', 'verificarYEnviarCorreos')
    .addSeparator()
    .addItem('🧹 Purgar estado/logs', 'purgarSistema')
    .addToUi();
}`,
'onOpen'
);

// 2) Bloque nuevo para ver codigo/estado
insertAfter(
`function configurarSistema() {
  SpreadsheetApp.getUi().alert('Configuracion deshabilitada. Este CRM usa API y password fijas en el codigo.\n' + FIRMA_APP);
}
`,
`

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
    a { color:#6ec9ff; }
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
        <button class="soft" onclick="abrirEditor()">🌐 Abrir editor Apps Script</button>
      </div>
      <pre id="code"></pre>
      <div id="msg2"></div>
    </div>
  </div>

  <script>
    let payload = null;
    let files = [];

    function getErr(e) {
      if (!e) return 'Error desconocido';
      if (typeof e === 'string') return e;
      return e.message || JSON.stringify(e);
    }

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

          document.getElementById('meta').innerHTML =
            'Script ID: ' + escapeHtml(payload.scriptId || '') +
            ' | Archivos: ' + files.length +
            ' | Cargado: ' + escapeHtml(payload.fetchedAt || '');

          render();
        })
        .withFailureHandler(function(e) {
          msg.textContent = '❌ ' + getErr(e);
        })
        .obtenerCodigoProyectoSeguro_(pass);
    }

    function render() {
      const sel = document.getElementById('sel');
      const idx = Number(sel.value || '0');
      const f = files[idx] || files[0] || {};
      document.getElementById('code').textContent = String(f.source || '');
    }

    function copiar() {
      const txt = document.getElementById('code').textContent || '';
      navigator.clipboard.writeText(txt).then(function(){
        alert('✅ Archivo copiado');
      }).catch(function(){
        alert('❌ No se pudo copiar');
      });
    }

    function abrirEditor() {
      const url = (payload && payload.editUrl) ? payload.editUrl : '';
      if (!url) {
        alert('No hay URL de editor disponible.');
        return;
      }
      window.open(url, '_blank');
    }

    function escapeHtml(t) {
      return String(t || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
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
    throw new Error(
      'No se pudo cargar el codigo automaticamente (HTTP ' + code + '). ' +
      'Revisa permisos de Apps Script API. Detalle: ' + body.substring(0, 220)
    );
  }

  const parsed = JSON.parse(resp.getContentText() || '{}');
  const files = (parsed.files || [])
    .map(function(f) {
      return {
        name: f.name || 'SIN_NOMBRE',
        type: f.type || 'DESCONOCIDO',
        source: String(f.source || '')
      };
    })
    .sort(function(a, b) {
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
        document.getElementById('meta').textContent =
          'Script ID: ' + (data.scriptId || '') +
          ' | Hora: ' + (data.now || '');
        document.getElementById('out').textContent = JSON.stringify(data, null, 2);
      }).withFailureHandler(function(e){
        document.getElementById('meta').textContent = 'Error';
        document.getElementById('out').textContent = '❌ ' + (e && e.message ? e.message : e);
      }).obtenerEstadoTecnico_();
    }

    function copiar() {
      const txt = document.getElementById('out').textContent || '';
      navigator.clipboard.writeText(txt).then(function(){
        alert('✅ Estado copiado');
      });
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
    'STOP_REQUESTED', 'IDX_EXISTING', 'IDX_NEW', 'IDX_MODEL'
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
`,
'insert viewers after configurarSistema'
);

// 3) Parada sin popup bloqueante
replaceOnce(
`function solicitarParada() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('STOP_REQUESTED', 'TRUE');
  logCRM_('🛑 Peticion de parada recibida.', 'warning');
  SpreadsheetApp.getUi().alert('🛑 Peticion de parada registrada.');
}`,
`function solicitarParada() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('STOP_REQUESTED', 'TRUE');
  logCRM_('🛑 Peticion de parada recibida.', 'warning');
  return 'OK';
}`,
'solicitarParada'
);

// 4) Etiqueta unica para relleno radar
s = s.replace(
`      'Radar pendiente de validacion manual',`,
`      'Radar pendiente de validacion manual [' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm') + ' #' + (rowsToInsert.length + 1) + ']',`
);

fs.writeFileSync(path, s, 'utf8');
console.log('PATCH_OK');
