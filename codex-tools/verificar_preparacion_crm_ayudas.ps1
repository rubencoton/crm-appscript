param(
  [string]$SpreadsheetId = '1LgZG2ObSCJzEQvrysu1NFFEvYlupLXVByDnIMCr-wYA',
  [string]$ProjectDir = '',
  [string]$ClasprcPath = '',
  [switch]$Json,
  [switch]$Strict
)

$ErrorActionPreference = 'Stop'

function Resolve-DefaultProjectDir {
  $base = Split-Path -Parent $PSScriptRoot
  return (Join-Path $base 'crm-ayudas-subvenciones')
}

function Resolve-DefaultClasprcPath {
  return (Join-Path $env:USERPROFILE '.clasprc.json')
}

function Add-Check {
  param(
    [System.Collections.Generic.List[object]]$Checks,
    [string]$Name,
    [bool]$Ok,
    [string]$Detail
  )
  $Checks.Add([PSCustomObject]@{
      check = $Name
      ok = $Ok
      detail = $Detail
    }) | Out-Null
}

function Parse-ApiError {
  param($ErrorRecord)
  if ($ErrorRecord.Exception -and $ErrorRecord.Exception.Response) {
    try {
      $sr = New-Object IO.StreamReader($ErrorRecord.Exception.Response.GetResponseStream())
      $txt = $sr.ReadToEnd()
      $sr.Close()
      if (-not [string]::IsNullOrWhiteSpace($txt)) { return $txt }
    }
    catch {}
  }
  return ($ErrorRecord.Exception.Message)
}

function Get-NodeExe {
  $cmd = Get-Command node.exe -ErrorAction SilentlyContinue
  if ($cmd -and $cmd.Source) { return $cmd.Source }
  $candidates = @(
    'C:\Program Files\nodejs\node.exe',
    'C:\Program Files (x86)\nodejs\node.exe',
    'C:\Progra~1\nodejs\node.exe'
  )
  foreach ($p in $candidates) {
    if (Test-Path -LiteralPath $p) { return $p }
  }
  throw 'No se encontro node.exe'
}

function Get-AccessTokenFromClasprc {
  param([string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) {
    throw "No existe .clasprc.json en $Path"
  }

  $cfg = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json

  # Formato nuevo de clasp (token + oauth2ClientSettings)
  if ($cfg.token -and $cfg.oauth2ClientSettings) {
    $rt = [string]$cfg.token.refresh_token
    $cid = [string]$cfg.oauth2ClientSettings.clientId
    $csec = [string]$cfg.oauth2ClientSettings.clientSecret
    if ($rt -and $cid -and $csec) {
      $resp = Invoke-RestMethod -Method Post -Uri 'https://oauth2.googleapis.com/token' -Body @{
        client_id = $cid
        client_secret = $csec
        refresh_token = $rt
        grant_type = 'refresh_token'
      }
      if (-not $resp.access_token) { throw 'Refresh OAuth sin access_token (formato nuevo)' }
      return [string]$resp.access_token
    }
  }

  # Formato antiguo (tokens.default)
  if ($cfg.tokens -and $cfg.tokens.default) {
    $tok = $cfg.tokens.default
    $rt = [string]$tok.refresh_token
    $cid = [string]$tok.client_id
    $csec = [string]$tok.client_secret
    if ($rt -and $cid -and $csec) {
      $resp = Invoke-RestMethod -Method Post -Uri 'https://oauth2.googleapis.com/token' -Body @{
        client_id = $cid
        client_secret = $csec
        refresh_token = $rt
        grant_type = 'refresh_token'
      }
      if (-not $resp.access_token) { throw 'Refresh OAuth sin access_token (formato antiguo)' }
      return [string]$resp.access_token
    }
  }

  throw 'No se encontraron credenciales OAuth reutilizables en .clasprc.json'
}

function Invoke-GApiGet {
  param(
    [string]$Url,
    [string]$AccessToken
  )
  return Invoke-RestMethod -Method Get -Uri $Url -Headers @{ Authorization = ('Bearer ' + $AccessToken) }
}

if (-not $ProjectDir) { $ProjectDir = Resolve-DefaultProjectDir }
if (-not $ClasprcPath) { $ClasprcPath = Resolve-DefaultClasprcPath }

$checks = New-Object 'System.Collections.Generic.List[object]'
$meta = [ordered]@{
  timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  projectDir = $ProjectDir
  spreadsheetId = $SpreadsheetId
  claspScriptId = $null
  scriptTitle = $null
  scriptParentId = $null
  spreadsheetTitle = $null
  parentSpreadsheetTitle = $null
  ready = $false
}

try {
  Add-Check $checks 'path.projectDir' (Test-Path -LiteralPath $ProjectDir) $ProjectDir
  if (-not (Test-Path -LiteralPath $ProjectDir)) { throw 'ProjectDir no existe' }

  $claspPath = Join-Path $ProjectDir '.clasp.json'
  Add-Check $checks 'path.claspJson' (Test-Path -LiteralPath $claspPath) $claspPath
  if (-not (Test-Path -LiteralPath $claspPath)) { throw '.clasp.json no existe' }

  $cl = Get-Content -LiteralPath $claspPath -Raw | ConvertFrom-Json
  $meta.claspScriptId = [string]$cl.scriptId
  Add-Check $checks 'clasp.scriptId' (-not [string]::IsNullOrWhiteSpace($meta.claspScriptId)) ('scriptId=' + $meta.claspScriptId)

  $nodeExe = Get-NodeExe
  $claspJs = Join-Path (Split-Path -Parent $ProjectDir) 'node_modules\@google\clasp\build\src\index.js'
  $statusOut = & $nodeExe $claspJs status 2>&1
  $statusOk = ($LASTEXITCODE -eq 0)
  $statusDetail = if ($statusOk) { 'OK' } else { ($statusOut -join ' | ') }
  Add-Check $checks 'clasp.status' $statusOk $statusDetail

  $access = Get-AccessTokenFromClasprc -Path $ClasprcPath
  Add-Check $checks 'oauth.refresh' (-not [string]::IsNullOrWhiteSpace($access)) 'token renovado'

  $script = Invoke-GApiGet -Url ('https://script.googleapis.com/v1/projects/' + $meta.claspScriptId) -AccessToken $access
  $meta.scriptTitle = [string]$script.title
  $meta.scriptParentId = [string]$script.parentId
  Add-Check $checks 'script.getProject' $true ('title=' + $meta.scriptTitle + '; parentId=' + $meta.scriptParentId)

  $sheet = Invoke-GApiGet -Url ('https://www.googleapis.com/drive/v3/files/' + $SpreadsheetId + '?fields=id,name,mimeType') -AccessToken $access
  $meta.spreadsheetTitle = [string]$sheet.name
  Add-Check $checks 'drive.targetSheet' $true ('title=' + $meta.spreadsheetTitle)

  if ($meta.scriptParentId) {
    $parent = Invoke-GApiGet -Url ('https://www.googleapis.com/drive/v3/files/' + $meta.scriptParentId + '?fields=id,name,mimeType') -AccessToken $access
    $meta.parentSpreadsheetTitle = [string]$parent.name
    Add-Check $checks 'drive.scriptParentSheet' $true ('title=' + $meta.parentSpreadsheetTitle)
  }
  else {
    Add-Check $checks 'drive.scriptParentSheet' $false 'script sin parentId (script no contenedor)'
  }

  $match = ($meta.scriptParentId -eq $SpreadsheetId)
  Add-Check $checks 'binding.parentMatchesTarget' $match ('parentId=' + $meta.scriptParentId + '; target=' + $SpreadsheetId)
}
catch {
  Add-Check $checks 'fatal' $false (Parse-ApiError $_)
}

$requiredChecks = @(
  'path.projectDir',
  'path.claspJson',
  'clasp.scriptId',
  'clasp.status',
  'oauth.refresh',
  'script.getProject',
  'drive.targetSheet',
  'drive.scriptParentSheet',
  'binding.parentMatchesTarget'
)

$map = @{}
foreach ($c in $checks) { $map[$c.check] = $c.ok }

$allOk = $true
foreach ($name in $requiredChecks) {
  if (-not $map.ContainsKey($name) -or -not $map[$name]) {
    $allOk = $false
    break
  }
}
$meta.ready = $allOk

$out = [ordered]@{
  meta = $meta
  checks = $checks
}

if ($Json) {
  $out | ConvertTo-Json -Depth 8
}
else {
  Write-Output '========================================'
  Write-Output ' VERIFICACION CRM AYUDAS (PREPARACION) '
  Write-Output '========================================'
  Write-Output ('ProjectDir: ' + $meta.projectDir)
  Write-Output ('SpreadsheetId objetivo: ' + $meta.spreadsheetId)
  Write-Output ('ScriptId clasp: ' + $meta.claspScriptId)
  Write-Output ('Script title: ' + $meta.scriptTitle)
  Write-Output ('Script parentId: ' + $meta.scriptParentId)
  Write-Output ('Hoja objetivo: ' + $meta.spreadsheetTitle)
  Write-Output ('Hoja parent del script: ' + $meta.parentSpreadsheetTitle)
  Write-Output ''
  foreach ($c in $checks) {
    $tag = if ($c.ok) { '[OK]' } else { '[FAIL]' }
    Write-Output ($tag + ' ' + $c.check + ' -> ' + $c.detail)
  }
  Write-Output ''
  Write-Output ('READY=' + $meta.ready)
}

if ($Strict -and -not $meta.ready) {
  exit 2
}
