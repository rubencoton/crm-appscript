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

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$crmPath = Join-Path $repoRoot "crm-ayudas-subvenciones"

Ensure-ToolPath
$git = Get-GitCmd
Set-Location $repoRoot

Write-Host "[1/5] clasp push (Festivales)"
$code = Invoke-Clasp -RepoRoot $repoRoot -ProjectPath $repoRoot -Action "push"
if ($code -ne 0) { throw "clasp push en Festivales fallo" }

Write-Host "[2/5] clasp push (CRM AYUDAS Y SUBVENCIONES)"
$code = Invoke-Clasp -RepoRoot $repoRoot -ProjectPath $crmPath -Action "push"
if ($code -ne 0) { throw "clasp push en CRM fallo" }

Write-Host "[3/5] git add"
& $git -C $repoRoot add -A
if ($LASTEXITCODE -ne 0) { throw "git add fallo" }

$changes = & $git -C $repoRoot diff --cached --name-only
if (-not $changes) {
  Write-Host "[4/5] No hay cambios para commit"
  Write-Host "[5/5] No se hace git push (nada nuevo)"
  Write-Host "OK sync_out completado (sin cambios de Git)"
  exit 0
}

$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$msg = "sync_out $stamp"

Write-Host "[4/5] git commit"
& $git -C $repoRoot commit -m $msg
if ($LASTEXITCODE -ne 0) { throw "git commit fallo" }

Write-Host "[5/5] git push"
try {
  & $git -C $repoRoot push origin main
  if ($LASTEXITCODE -ne 0) { throw "git push devolvio codigo $LASTEXITCODE" }
} catch {
  Write-Host "ERROR de Git push (sin forzar nada)."
  Write-Host "Si hay rechazo/conflicto remoto, primero haz 'git pull --ff-only' y revisa estado."
  throw
}

Write-Host "OK sync_out completado"
