param(
  [switch]$DryRun,
  [switch]$SkipCrmProject,
  [switch]$SkipGitPush,
  [string]$CommitMessage = ""
)

$ErrorActionPreference = "Stop"

function Ensure-ToolPath {
  $candidates = @(
    "C:\Program Files\nodejs",
    "C:\Program Files\Git\cmd",
    "C:\Program Files\Git\mingw64\bin"
  )
  foreach ($dir in $candidates) {
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

function Invoke-Clasp {
  param(
    [string]$RepoRoot,
    [string]$ProjectPath,
    [string]$Action
  )

  $localClasp = Join-Path $RepoRoot "node_modules\\.bin\\clasp.cmd"
  if (Test-Path $localClasp) {
    & $localClasp -P $ProjectPath $Action
    return $LASTEXITCODE
  }

  $claspCmd = Get-Command clasp.cmd -ErrorAction SilentlyContinue
  if ($claspCmd) {
    & $claspCmd.Source -P $ProjectPath $Action
    return $LASTEXITCODE
  }

  $npx = "C:\Program Files\nodejs\npx.cmd"
  if (Test-Path $npx) {
    & $npx clasp -P $ProjectPath $Action
    return $LASTEXITCODE
  }

  throw "No se encontro clasp.cmd ni npx.cmd"
}

function Run-OrDry {
  param(
    [string]$Title,
    [scriptblock]$Action
  )
  Write-Host $Title
  if ($DryRun) {
    Write-Host "[DRYRUN] $Title"
    return 0
  }
  & $Action
  return $LASTEXITCODE
}

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$crmPath = Join-Path $repoRoot "crm-ayudas-subvenciones"

Ensure-ToolPath
$git = Get-GitCmd
Set-Location $repoRoot

$code = Run-OrDry -Title "[1/5] clasp push (Festivales)" -Action {
  Invoke-Clasp -RepoRoot $repoRoot -ProjectPath $repoRoot -Action "push"
}
if ($code -ne 0) { throw "clasp push en Festivales fallo" }

if ($SkipCrmProject) {
  Write-Host "[2/5] omitido por -SkipCrmProject"
} else {
  if (-not (Test-Path -LiteralPath (Join-Path $crmPath ".clasp.json"))) {
    throw "No existe .clasp.json en proyecto CRM secundario: $crmPath"
  }
  $code = Run-OrDry -Title "[2/5] clasp push (CRM AYUDAS Y SUBVENCIONES)" -Action {
    Invoke-Clasp -RepoRoot $repoRoot -ProjectPath $crmPath -Action "push"
  }
  if ($code -ne 0) { throw "clasp push en CRM fallo" }
}

$code = Run-OrDry -Title "[3/5] git add -A" -Action {
  & $git -c "safe.directory=$repoRoot" -C $repoRoot add -A
}
if ($code -ne 0) { throw "git add fallo" }

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

$code = Run-OrDry -Title "[4/5] git commit" -Action {
  & $git -c "safe.directory=$repoRoot" -C $repoRoot commit -m $msg
}
if ($code -ne 0) { throw "git commit fallo" }

if ($SkipGitPush) {
  Write-Host "[5/5] omitido por -SkipGitPush"
  Write-Host "OK sync_out completado"
  exit 0
}

$code = Run-OrDry -Title "[5/5] git push" -Action {
  & $git -c "safe.directory=$repoRoot" -C $repoRoot push origin main
}
if ($code -ne 0) {
  Write-Host "ERROR de Git push (sin forzar nada)."
  Write-Host "Si hay rechazo/conflicto remoto, primero haz 'git pull --ff-only' y revisa estado."
  throw "git push devolvio codigo $code"
}

Write-Host "OK sync_out completado"
