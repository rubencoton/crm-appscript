param(
  [switch]$DryRun,
  [switch]$SkipCrmProject,
  [switch]$SkipGitPush,
  [switch]$AllowBulkUntracked,
  [string]$CommitMessage = ""
)

$ErrorActionPreference = "Stop"

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

function Get-ClaspCmd {
  param([string]$RepoRoot)

  $localClasp = Join-Path $RepoRoot "node_modules\.bin\clasp.cmd"
  if (Test-Path $localClasp) { return $localClasp }

  $claspCmd = Get-Command clasp.cmd -ErrorAction SilentlyContinue
  if ($claspCmd) { return $claspCmd.Source }

  $directWinGet = Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Packages\OpenJS.NodeJS.LTS_Microsoft.Winget.Source_8wekyb3d8bbwe\node-v24.14.0-win-x64\clasp.cmd"
  if (Test-Path $directWinGet) { return $directWinGet }

  $wingetNode = Get-WinGetNodeDir
  if ($wingetNode) {
    $wingetClasp = Join-Path $wingetNode "clasp.cmd"
    if (Test-Path $wingetClasp) { return $wingetClasp }
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

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$crmPath = Join-Path $repoRoot "crm-ayudas-subvenciones"

Ensure-ToolPath
$git = Get-GitCmd
Set-Location $repoRoot

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

Invoke-Step -Title "[1/5] clasp push (Festivales)" -ErrorMessage "clasp push en Festivales fallo" -Action {
  Invoke-Clasp -RepoRoot $repoRoot -ProjectPath $repoRoot -Action "push"
}

if ($SkipCrmProject) {
  Write-Host "[2/5] omitido por -SkipCrmProject"
} else {
  if (-not (Test-Path -LiteralPath (Join-Path $crmPath ".clasp.json"))) {
    throw "No existe .clasp.json en proyecto CRM secundario: $crmPath"
  }

  Invoke-Step -Title "[2/5] clasp push (CRM AYUDAS Y SUBVENCIONES)" -ErrorMessage "clasp push en CRM fallo" -Action {
    Invoke-Clasp -RepoRoot $repoRoot -ProjectPath $crmPath -Action "push"
  }
}

Invoke-Step -Title "[3/5] git add -A" -ErrorMessage "git add fallo" -Action {
  & $git -c "safe.directory=$repoRoot" -C $repoRoot add -A
}

if ($DryRun) {
  Write-Host "[4/5] commit omitido (DryRun)"
  Write-Host "[5/5] push omitido (DryRun)"
  Write-Host "OK sync_out completado (DryRun)"
  exit 0
}

$changes = & $git -c "safe.directory=$repoRoot" -C $repoRoot diff --cached --name-only
if (-not $changes) {
  Write-Host "[4/5] No hay cambios para commit"
  Write-Host "[5/5] No se hace git push (nada nuevo)"
  Write-Host "OK sync_out completado (sin cambios de Git)"
  exit 0
}

$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$msg = $CommitMessage
if ([string]::IsNullOrWhiteSpace($msg)) {
  $msg = "sync_out $stamp"
}

Invoke-Step -Title "[4/5] git commit" -ErrorMessage "git commit fallo" -Action {
  & $git -c "safe.directory=$repoRoot" -C $repoRoot commit -m $msg
}

if ($SkipGitPush) {
  Write-Host "[5/5] omitido por -SkipGitPush"
  Write-Host "OK sync_out completado"
  exit 0
}

try {
  Invoke-Step -Title "[5/5] git push" -ErrorMessage "git push fallo" -Action {
    & $git -c "safe.directory=$repoRoot" -C $repoRoot push origin main
  }
} catch {
  Write-Host "ERROR de Git push (sin forzar nada)."
  Write-Host "Si hay rechazo/conflicto remoto, primero haz 'git pull --ff-only' y revisa estado."
  throw
}

Write-Host "OK sync_out completado"
