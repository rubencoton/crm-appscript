param(
  [string]$SpreadsheetId = '1LgZG2ObSCJzEQvrysu1NFFEvYlupLXVByDnIMCr-wYA',
  [string]$ProjectDir = '',
  [string]$ClasprcPath = '',
  [string]$WorkspaceRoot = '',
  [string]$KnownScriptId = '',
  [int]$MaxDriveCandidates = 500,
  [switch]$AutoFix,
  [switch]$Pull,
  [switch]$Json,
  [switch]$Strict
)

$ErrorActionPreference = 'Stop'

function Resolve-DefaultProjectDir {
  $base = Split-Path -Parent $PSScriptRoot
  return (Join-Path $base 'crm-ayudas-subvenciones')
}

function Resolve-DefaultWorkspaceRoot {
  return (Split-Path -Parent $PSScriptRoot)
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

function Get-ScriptProjectMeta {
  param(
    [string]$ScriptId,
    [string]$AccessToken
  )
  if ([string]::IsNullOrWhiteSpace($ScriptId)) { return $null }
  try {
    $m = Invoke-GApiGet -Url ('https://script.googleapis.com/v1/projects/' + $ScriptId) -AccessToken $AccessToken
    return [PSCustomObject]@{
      ok = $true
      scriptId = $ScriptId
      title = [string]$m.title
      parentId = [string]$m.parentId
      raw = $m
      error = ''
    }
  }
  catch {
    return [PSCustomObject]@{
      ok = $false
      scriptId = $ScriptId
      title = ''
      parentId = ''
      raw = $null
      error = (Parse-ApiError $_)
    }
  }
}

function Get-DriveScriptIds {
  param(
    [string]$AccessToken,
    [int]$MaxItems = 500
  )
  $out = New-Object 'System.Collections.Generic.List[string]'
  $pageToken = $null
  do {
    $url = 'https://www.googleapis.com/drive/v3/files?q=' + [uri]::EscapeDataString("mimeType='application/vnd.google-apps.script' and trashed=false") + '&fields=files(id),nextPageToken&pageSize=200'
    if ($pageToken) { $url += '&pageToken=' + [uri]::EscapeDataString($pageToken) }
    $resp = Invoke-GApiGet -Url $url -AccessToken $AccessToken
    foreach ($f in @($resp.files)) {
      if ($f.id -and -not $out.Contains([string]$f.id)) {
        $out.Add([string]$f.id) | Out-Null
        if ($out.Count -ge $MaxItems) { break }
      }
    }
    if ($out.Count -ge $MaxItems) { break }
    $pageToken = [string]$resp.nextPageToken
  } while ($pageToken)
  return ,$out.ToArray()
}

function Get-LocalScriptIdsFromClasp {
  param([string]$RootPath)
  $ids = New-Object 'System.Collections.Generic.List[string]'
  if (-not (Test-Path -LiteralPath $RootPath)) { return ,$ids.ToArray() }
  $files = Get-ChildItem -Path $RootPath -Recurse -File -Filter '.clasp.json' -ErrorAction SilentlyContinue
  foreach ($f in $files) {
    try {
      $c = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json
      $id = [string]$c.scriptId
      if ($id -and -not $ids.Contains($id)) { $ids.Add($id) | Out-Null }
    }
    catch {}
  }
  return ,$ids.ToArray()
}

function Set-LocalClaspScriptId {
  param(
    [string]$ClaspPath,
    [string]$ScriptId
  )
  $now = Get-Date -Format 'yyyyMMdd_HHmmss'
  $bak = $ClaspPath + '.bak-' + $now
  Copy-Item -LiteralPath $ClaspPath -Destination $bak -Force
  $j = Get-Content -LiteralPath $ClaspPath -Raw | ConvertFrom-Json
  $j.scriptId = $ScriptId
  ($j | ConvertTo-Json -Depth 20) + "`n" | Set-Content -LiteralPath $ClaspPath -NoNewline
  return $bak
}

function Invoke-ClaspCommand {
  param(
    [string]$ProjectDir,
    [string]$Args
  )
  $node = Get-NodeExe
  $claspJs = Join-Path (Split-Path -Parent $ProjectDir) 'node_modules\@google\clasp\build\src\index.js'
  if (-not (Test-Path -LiteralPath $claspJs)) {
    throw "No se encontro clasp local en $claspJs"
  }
  $argList = @($claspJs) + (@($Args -split '\s+') | Where-Object { $_ })
  Push-Location $ProjectDir
  try {
    $out = & $node @argList 2>&1
    return [PSCustomObject]@{
      exitCode = $LASTEXITCODE
      output = @($out)
    }
  }
  finally {
    Pop-Location
  }
}

if (-not $ProjectDir) { $ProjectDir = Resolve-DefaultProjectDir }
if (-not $WorkspaceRoot) { $WorkspaceRoot = Resolve-DefaultWorkspaceRoot }
if (-not $ClasprcPath) { $ClasprcPath = Resolve-DefaultClasprcPath }

$checks = New-Object 'System.Collections.Generic.List[object]'
$discovery = [ordered]@{
  driveCandidates = @()
  localClaspCandidates = @()
  inspectedScripts = 0
  matchedScriptIds = @()
}
$action = [ordered]@{
  autoFixApplied = $false
  pullApplied = $false
  backupPath = ''
  oldScriptId = ''
  newScriptId = ''
}
$meta = [ordered]@{
  timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  projectDir = $ProjectDir
  workspaceRoot = $WorkspaceRoot
  spreadsheetId = $SpreadsheetId
  spreadsheetTitle = ''
  localScriptId = ''
  localScriptTitle = ''
  localScriptParentId = ''
  localParentMatchesTarget = $false
  discoveredScriptId = ''
  discoveredScriptTitle = ''
  discoveredScriptParentId = ''
  ready = $false
}

try {
  Add-Check $checks 'path.projectDir' (Test-Path -LiteralPath $ProjectDir) $ProjectDir
  if (-not (Test-Path -LiteralPath $ProjectDir)) { throw 'ProjectDir no existe' }

  $claspPath = Join-Path $ProjectDir '.clasp.json'
  Add-Check $checks 'path.claspJson' (Test-Path -LiteralPath $claspPath) $claspPath
  if (-not (Test-Path -LiteralPath $claspPath)) { throw '.clasp.json no existe' }

  $cl = Get-Content -LiteralPath $claspPath -Raw | ConvertFrom-Json
  $meta.localScriptId = [string]$cl.scriptId
  $action.oldScriptId = $meta.localScriptId
  Add-Check $checks 'clasp.scriptId.present' (-not [string]::IsNullOrWhiteSpace($meta.localScriptId)) ('scriptId=' + $meta.localScriptId)

  $access = Get-AccessTokenFromClasprc -Path $ClasprcPath
  Add-Check $checks 'oauth.refresh' (-not [string]::IsNullOrWhiteSpace($access)) 'token renovado'

  $sheet = Invoke-GApiGet -Url ('https://www.googleapis.com/drive/v3/files/' + $SpreadsheetId + '?fields=id,name,mimeType') -AccessToken $access
  $meta.spreadsheetTitle = [string]$sheet.name
  Add-Check $checks 'drive.targetSheet' ($sheet.mimeType -eq 'application/vnd.google-apps.spreadsheet') ('title=' + $meta.spreadsheetTitle)

  $localMeta = Get-ScriptProjectMeta -ScriptId $meta.localScriptId -AccessToken $access
  $localMetaDetail = if ($localMeta.ok) { 'title=' + $localMeta.title + '; parentId=' + $localMeta.parentId } else { $localMeta.error }
  Add-Check $checks 'script.local.getProject' $localMeta.ok $localMetaDetail
  if ($localMeta.ok) {
    $meta.localScriptTitle = $localMeta.title
    $meta.localScriptParentId = $localMeta.parentId
    $meta.localParentMatchesTarget = ($localMeta.parentId -eq $SpreadsheetId)
  }

  Add-Check $checks 'binding.localParentMatchesTarget' $meta.localParentMatchesTarget ('parentId=' + $meta.localScriptParentId + '; target=' + $SpreadsheetId)

  if (-not $meta.localParentMatchesTarget) {
    $driveIds = Get-DriveScriptIds -AccessToken $access -MaxItems $MaxDriveCandidates
    $discovery.driveCandidates = @($driveIds)
    Add-Check $checks 'discover.driveCandidates' ($driveIds.Count -gt 0) ('count=' + $driveIds.Count)

    $localIds = Get-LocalScriptIdsFromClasp -RootPath $WorkspaceRoot
    $discovery.localClaspCandidates = @($localIds)
    Add-Check $checks 'discover.localClaspCandidates' ($localIds.Count -gt 0) ('count=' + $localIds.Count)

    $candidateIds = New-Object 'System.Collections.Generic.List[string]'
    foreach ($id in $driveIds) { if ($id -and -not $candidateIds.Contains($id)) { $candidateIds.Add($id) | Out-Null } }
    foreach ($id in $localIds) { if ($id -and -not $candidateIds.Contains($id)) { $candidateIds.Add($id) | Out-Null } }
    if ($KnownScriptId -and -not $candidateIds.Contains($KnownScriptId)) { $candidateIds.Add($KnownScriptId) | Out-Null }

    $found = New-Object 'System.Collections.Generic.List[object]'
    foreach ($sid in $candidateIds) {
      $m = Get-ScriptProjectMeta -ScriptId $sid -AccessToken $access
      $discovery.inspectedScripts++
      if ($m.ok -and $m.parentId -eq $SpreadsheetId) {
        $found.Add($m) | Out-Null
      }
    }

    $discovery.matchedScriptIds = @($found | ForEach-Object { $_.scriptId })
    Add-Check $checks 'discover.boundScriptFound' ($found.Count -gt 0) ('matches=' + $found.Count + '; inspected=' + $discovery.inspectedScripts)

    if ($found.Count -gt 0) {
      $pick = $found[0]
      $meta.discoveredScriptId = [string]$pick.scriptId
      $meta.discoveredScriptTitle = [string]$pick.title
      $meta.discoveredScriptParentId = [string]$pick.parentId
    }
  }

  if ($AutoFix -and $meta.discoveredScriptId) {
    if ($meta.discoveredScriptId -ne $meta.localScriptId) {
      $bak = Set-LocalClaspScriptId -ClaspPath $claspPath -ScriptId $meta.discoveredScriptId
      $action.autoFixApplied = $true
      $action.backupPath = $bak
      $action.newScriptId = $meta.discoveredScriptId
      Add-Check $checks 'fix.claspScriptIdUpdated' $true ('newScriptId=' + $action.newScriptId)
    }
    else {
      Add-Check $checks 'fix.claspScriptIdUpdated' $true 'ya estaba actualizado'
    }
  }
  elseif ($AutoFix -and -not $meta.discoveredScriptId) {
    Add-Check $checks 'fix.claspScriptIdUpdated' $false 'AutoFix activo pero no se detecto scriptId vinculado'
  }

  if ($Pull) {
    $st = Invoke-ClaspCommand -ProjectDir $ProjectDir -Args 'status'
    Add-Check $checks 'clasp.status' ($st.exitCode -eq 0) (($st.output -join ' | '))
    if ($st.exitCode -eq 0) {
      $pl = Invoke-ClaspCommand -ProjectDir $ProjectDir -Args 'pull'
      $ok = ($pl.exitCode -eq 0)
      Add-Check $checks 'clasp.pull' $ok (($pl.output -join ' | '))
      $action.pullApplied = $ok
    }
  }

  $cl2 = Get-Content -LiteralPath $claspPath -Raw | ConvertFrom-Json
  $currentScriptId = [string]$cl2.scriptId
  $currentMeta = Get-ScriptProjectMeta -ScriptId $currentScriptId -AccessToken $access
  $matchNow = ($currentMeta.ok -and $currentMeta.parentId -eq $SpreadsheetId)
  $finalDetail = if ($currentMeta.ok) { 'parentId=' + $currentMeta.parentId + '; target=' + $SpreadsheetId } else { $currentMeta.error }
  Add-Check $checks 'binding.finalParentMatchesTarget' $matchNow $finalDetail

  $meta.localScriptId = $currentScriptId
  if ($currentMeta.ok) {
    $meta.localScriptTitle = $currentMeta.title
    $meta.localScriptParentId = $currentMeta.parentId
  }
  $meta.localParentMatchesTarget = $matchNow
}
catch {
  Add-Check $checks 'fatal' $false (Parse-ApiError $_)
}

$required = @(
  'path.projectDir',
  'path.claspJson',
  'clasp.scriptId.present',
  'oauth.refresh',
  'drive.targetSheet',
  'script.local.getProject',
  'binding.finalParentMatchesTarget'
)

$map = @{}
foreach ($c in $checks) { $map[$c.check] = $c.ok }
$allOk = $true
foreach ($r in $required) {
  if (-not $map.ContainsKey($r) -or -not $map[$r]) { $allOk = $false; break }
}
$meta.ready = $allOk

$result = [ordered]@{
  meta = $meta
  discovery = $discovery
  action = $action
  checks = $checks
}

if ($Json) {
  $result | ConvertTo-Json -Depth 12
}
else {
  Write-Output '==========================================='
  Write-Output ' PREPARACION CRM AYUDAS (AUTO-CHECK READY) '
  Write-Output '==========================================='
  Write-Output ('ProjectDir: ' + $meta.projectDir)
  Write-Output ('Spreadsheet objetivo: ' + $meta.spreadsheetId + ' | ' + $meta.spreadsheetTitle)
  Write-Output ('ScriptId local final: ' + $meta.localScriptId)
  Write-Output ('Script title: ' + $meta.localScriptTitle)
  Write-Output ('Script parentId: ' + $meta.localScriptParentId)
  if ($meta.discoveredScriptId) {
    Write-Output ('ScriptId detectado por discovery: ' + $meta.discoveredScriptId + ' | ' + $meta.discoveredScriptTitle)
  }
  else {
    Write-Output 'ScriptId detectado por discovery: (ninguno)'
  }
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
