$ErrorActionPreference = "Stop"

function Get-GitCmd {
  $gitCmd = Get-Command git.exe -ErrorAction SilentlyContinue
  if ($gitCmd) { return $gitCmd.Source }
  $fallback = "C:\Program Files\Git\cmd\git.exe"
  if (Test-Path $fallback) { return $fallback }
  throw "No se encontro git.exe"
}

function Get-ClaspCmd {
  $claspCmd = Get-Command clasp.cmd -ErrorAction SilentlyContinue
  if ($claspCmd) { return $claspCmd.Source }
  $fallback = "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\OpenJS.NodeJS.LTS_Microsoft.Winget.Source_8wekyb3d8bbwe\node-v24.14.0-win-x64\clasp.cmd"
  if (Test-Path $fallback) { return $fallback }
  throw "No se encontro clasp.cmd"
}

$git = Get-GitCmd
$clasp = Get-ClaspCmd
$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$crmPath = Join-Path $repoRoot "crm-ayudas-subvenciones"
Set-Location $repoRoot

Write-Host "[1/5] clasp push (Festivales)"
& $clasp -P $repoRoot push
if ($LASTEXITCODE -ne 0) { throw "clasp push en Festivales fallo" }

Write-Host "[2/5] clasp push (CRM AYUDAS Y SUBVENCIONES)"
& $clasp -P $crmPath push
if ($LASTEXITCODE -ne 0) { throw "clasp push en CRM fallo" }

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
  & $git -C $repoRoot push
  if ($LASTEXITCODE -ne 0) { throw "git push devolvio codigo $LASTEXITCODE" }
} catch {
  Write-Host "ERROR de Git push (sin forzar nada)."
  Write-Host "Si hay rechazo/conflicto remoto, primero haz 'git pull --rebase' y resuelve manualmente."
  throw
}

Write-Host "OK sync_out completado"
