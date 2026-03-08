$p='C:\Users\elrub\Desktop\CARPETA CODEX\02_APPS_SCRIPT_FUENTES\crm_ayudas_prod\Code.gs'
$s=[IO.File]::ReadAllText($p,[Text.Encoding]::UTF8)
$s=$s.Replace("'5) Enlace bases actuales en link1 si existe.',","'5) Prioriza link1 como bases actuales, link2 historico y link3 complementario.',")
$s=$s.Replace("'7) Cuando no exista un dato, usa `"No publicado`".'","'7) Cuando no exista un dato, usa `"' + CRM_NO_DATA + '`".'")
$s=$s.Replace("'Busca 5 concursos o ayudas musicales vigentes en Espana.'","'Busca 5 concursos o ayudas musicales vigentes para bandas de cualquier parte del mundo.'")
[IO.File]::WriteAllText($p,$s,[Text.Encoding]::UTF8)
