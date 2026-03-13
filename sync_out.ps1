param(
  [switch]$DryRun,
  [switch]$SkipCrmProject,
  [switch]$SkipGitPush,
  [switch]$AllowBulkUntracked,
  [string]$CommitMessage = ""
)

$ErrorActionPreference = "Stop"

$ExpectedPrimaryScriptId = "1OGuPezQ26BFvaLRiy-IYIotGpmVu_Z_b9Mi8tCiprIz8zB4DgqmMc5Ea"
$ExpectedSpreadsheetId = "1kRrdCwd0n6FwVp-8rKP3gEeYT-TBXc1JL8sa_xzp1IM"
$ExpectedOrigin = "https://github.com/rubencoton/crm-appscript.git"

function Get-WinGetNodeDir {
  $packagesRoot = Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Packages"
  if (-not (Test-Path $packagesRoot)) { return $null }

  $pkg = Get-ChildItem -Path $packagesRoot -Directory -Filter "OpenJS.NodeJS.LTS*" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
  if (-not $pkg) { return $null }

  $nodeDir = Get-ChildItem -Path $pkg.FullName -Directory -Filter "node-v*-win-x64" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

  if ($nodeDir) { return $nodeDir.FullName }
  return $null
}

function Ensure-ToolPath {
  $wingetNode = Get-WinGetNodeDir
  $candidates = @(
    "C:\Program Files\nodejs",
    "C:\Program Files\Git\cmd",
    "C:\Program Files\Git\mingw64\bin",
    "$env:LOCALAPPDATA\Microsoft\WinGet\Links",
    $wingetNode
  )
  foreach ($dir in $candidates) {
    if ([string]::IsNullOrWhiteSpace($dir)) { continue }
    if (-not (Test-Path $dir)) { continue }
    if ($env:Path -notlike "*$dir*") {
      $env:Path = "$dir;$env:Path"
    }
  }
}

function Get-GitCmd {
  $gitCmd = Get-Command git.exe -ErrorAction SilentlyContinue
  if ($gitCmd) { return $gitCmd.Source }

  $fallback = "C:\Program Files\Git\cmd\git.exe"
  if (Test-Path $fallback) { return $fallback }

  throw "No se encontro git.exe"
}

function Test-CommandExecutable {
  param(
    [string]$CommandPath,
    [string[]]$CommandArgs = @("--version")
  )

  if ([string]::IsNullOrWhiteSpace($CommandPath)) { return $false }
  try {
    & $CommandPath @CommandArgs *> $null
    return ($LASTEXITCODE -eq 0)
  } catch {
    return $false
  }
}

function Get-ClaspCmd {
  param([string]$RepoRoot)

  $localClasp = Join-Path $RepoRoot "node_modules\.bin\clasp.cmd"
  $roamingClasp = Join-Path $env:APPDATA "npm\clasp.cmd"

  $directWinGet = Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Packages\OpenJS.NodeJS.LTS_Microsoft.Winget.Source_8wekyb3d8bbwe\node-v24.14.0-win-x64\clasp.cmd"
  $claspCmd = Get-Command clasp.cmd -ErrorAction SilentlyContinue
  $claspFromPath = $null
  if ($claspCmd) { $claspFromPath = $claspCmd.Source }
  $wingetNode = Get-WinGetNodeDir
  $wingetClasp = $null
  if ($wingetNode) {
    $wingetClasp = Join-Path $wingetNode "clasp.cmd"
  }

  $candidates = @(
    $claspFromPath,
    $roamingClasp,
    $localClasp,
    $directWinGet,
    $wingetClasp
  ) | Select-Object -Unique

  foreach ($candidate in $candidates) {
    if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
    if (-not (Test-Path $candidate)) { continue }
    return $candidate
  }

  return $null
}

function Get-NpxCmd {
  $npxCmd = Get-Command npx.cmd -ErrorAction SilentlyContinue
  if ($npxCmd) { return $npxCmd.Source }

  $npxProgramFiles = "C:\Program Files\nodejs\npx.cmd"
  if (Test-Path $npxProgramFiles) { return $npxProgramFiles }

  $wingetNode = Get-WinGetNodeDir
  if ($wingetNode) {
    $npxWinGet = Join-Path $wingetNode "npx.cmd"
    if (Test-Path $npxWinGet) { return $npxWinGet }
  }

  return $null
}

function Invoke-Clasp {
  param(
    [string]$RepoRoot,
    [string]$ProjectPath,
    [string]$Action
  )

  $clasp = Get-ClaspCmd -RepoRoot $RepoRoot
  if ($clasp) {
    & $clasp -P $ProjectPath $Action
    return
  }

  $npx = Get-NpxCmd
  if ($npx) {
    & $npx clasp -P $ProjectPath $Action
    return
  }

  throw "No se encontro clasp.cmd ni npx.cmd"
}

function Invoke-Step {
  param(
    [string]$Title,
    [scriptblock]$Action,
    [string]$ErrorMessage
  )

  Write-Host $Title
  if ($DryRun) {
    Write-Host "[DRYRUN] $Title"
    return
  }

  & $Action
  if ($LASTEXITCODE -ne 0) {
    throw "$ErrorMessage (codigo $LASTEXITCODE)"
  }
}

function Invoke-GitPushWithRetry {
  param(
    [string]$Git,
    [string]$RepoRoot
  )

  Write-Host "[4/4] git push (intento 1)"
  & $Git -c "safe.directory=$RepoRoot" -C $RepoRoot push origin main
  if ($LASTEXITCODE -eq 0) {
    return
  }

  Write-Host "WARN: push rechazado; intento de auto-integracion con pull --rebase (sin forzar)."
  & $Git -c "safe.directory=$RepoRoot" -C $RepoRoot pull --rebase origin main
  if ($LASTEXITCODE -ne 0) {
    throw "git pull --rebase fallo durante auto-recuperacion (codigo $LASTEXITCODE). Resolver manualmente."
  }

  Write-Host "[4/4] git push (intento 2)"
  & $Git -c "safe.directory=$RepoRoot" -C $RepoRoot push origin main
  if ($LASTEXITCODE -ne 0) {
    throw "git push fallo tras reintento (codigo $LASTEXITCODE)."
  }
}

function Assert-RepoScope {
  param(
    [string]$RepoRoot,
    [string]$Git
  )

  $claspPath = Join-Path $RepoRoot ".clasp.json"
  if (-not (Test-Path -LiteralPath $claspPath)) {
    throw "Falta .clasp.json en repo principal."
  }

  $claspJson = Get-Content $claspPath -Raw | ConvertFrom-Json
  $currentScriptId = [string]$claspJson.scriptId
  if ($currentScriptId -ne $ExpectedPrimaryScriptId) {
    throw "Scope bloqueado: scriptId actual '$currentScriptId' no coincide con '$ExpectedPrimaryScriptId' (hoja $ExpectedSpreadsheetId)."
  }

  $origin = & $Git -c "safe.directory=$RepoRoot" -C $RepoRoot remote get-url origin
  if ([string]$origin -ne $ExpectedOrigin) {
    throw "Scope bloqueado: origin actual '$origin' no coincide con '$ExpectedOrigin'."
  }
}

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

Ensure-ToolPath
$git = Get-GitCmd
Set-Location $repoRoot

Assert-RepoScope -RepoRoot $repoRoot -Git $git
Write-Host "Scope OK | sheet=$ExpectedSpreadsheetId | script=$ExpectedPrimaryScriptId"
if ($SkipCrmProject) {
  Write-Host "Nota: -SkipCrmProject recibido. Modo hoja unica ya lo omite por defecto."
}

if (-not $DryRun) {
  $untracked = & $git -c "safe.directory=$repoRoot" -C $repoRoot status --porcelain | Where-Object { $_ -like "?? *" }

  $hasWorkspaceBackup = $false
  foreach ($line in $untracked) {
    if ($line -match "_full_workspace_backup") {
      $hasWorkspaceBackup = $true
      break
    }
  }

  if ($hasWorkspaceBackup) {
    throw "Detectado _full_workspace_backup en no-trackeados. Revisa .gitignore antes de sync_out."
  }

  if (($untracked.Count -gt 40) -and (-not $AllowBulkUntracked)) {
    throw "Demasiados no-trackeados ($($untracked.Count)). Usa -AllowBulkUntracked si es intencional."
  }
}

Invoke-Step -Title "[1/4] clasp push (CRM: FESTIVALES)" -ErrorMessage "clasp push en CRM FESTIVALES fallo" -Action {
  Invoke-Clasp -RepoRoot $repoRoot -ProjectPath $repoRoot -Action "push"
}

Invoke-Step -Title "[2/4] git add -A" -ErrorMessage "git add fallo" -Action {
  & $git -c "safe.directory=$repoRoot" -C $repoRoot add -A
}

if ($DryRun) {
  Write-Host "[3/4] commit omitido (DryRun)"
  Write-Host "[4/4] push omitido (DryRun)"
  Write-Host "OK sync_out completado (DryRun, modo hoja unica)"
  exit 0
}

$changes = & $git -c "safe.directory=$repoRoot" -C $repoRoot diff --cached --name-only
if (-not $changes) {
  Write-Host "[3/4] No hay cambios para commit"
  Write-Host "[4/4] No se hace git push (nada nuevo)"
  Write-Host "OK sync_out completado (sin cambios de Git, modo hoja unica)"
  exit 0
}

$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$msg = $CommitMessage
if ([string]::IsNullOrWhiteSpace($msg)) {
  $msg = "sync_out $stamp"
}

Invoke-Step -Title "[3/4] git commit" -ErrorMessage "git commit fallo" -Action {
  & $git -c "safe.directory=$repoRoot" -C $repoRoot commit -m $msg
}

if ($SkipGitPush) {
  Write-Host "[4/4] omitido por -SkipGitPush"
  Write-Host "OK sync_out completado (modo hoja unica)"
  exit 0
}

try {
  Invoke-GitPushWithRetry -Git $git -RepoRoot $repoRoot
} catch {
  Write-Host "ERROR de Git push (sin forzar nada)."
  Write-Host "Si hay conflicto real, resolver manualmente o ejecutar 'git rebase --abort'."
  throw
}

Write-Host "OK sync_out completado (modo hoja unica)"

