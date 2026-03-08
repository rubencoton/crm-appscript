import re
from pathlib import Path

path = Path(r'C:\Users\elrub\Desktop\CARPETA CODEX\crm_apps_script_v2.gs')
text = path.read_text(encoding='utf-8')

text = re.sub(
    r"function verificarYEnviarCorreos\(\) \{[\s\S]*?\n\}\n\nfunction enviarBoletin_\(\)",
    """function verificarYEnviarCorreos() {
  mostrarDialogoBoletinSeguro_();
}

function mostrarDialogoBoletinSeguro_() {
  const html = HtmlService.createHtmlOutput(`<!doctype html>
<html>
<head>
  <meta charset=\"utf-8\">
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
  <div class=\"card\">
    <h3>📨 Enviar Boletin</h3>
    <p>Introduce password (oculta) para confirmar el envio.</p>
    <input id=\"pwd\" type=\"password\" placeholder=\"Password del CRM\">
    <div class=\"row\">
      <button class=\"secondary\" onclick=\"google.script.host.close()\">Cancelar</button>
      <button class=\"primary\" onclick=\"enviar()\">🚀 Enviar</button>
    </div>
    <div id=\"status\"></div>
  </div>
  <script>
    function enviar() {
      if (!confirm('Se enviara el boletin a todas las bandas con email valido usando solo concursos ABIERTA. ¿Continuar?')) return;
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
  SpreadsheetApp.getUi().showModalDialog(html, '📨 Seguridad de envio');
}

function enviarBoletinSeguroDesdePanel_(pass) {
  if (!validarPasswordServidor(pass)) {
    throw new Error('Password incorrecta.');
  }
  enviarBoletin_();
  return '✅ Envio ejecutado. Revisa el aviso final del sistema.';
}

function enviarBoletin_()""",
    text,
    count=1
)

path.write_text(text, encoding='utf-8')
print('patched boletin')
