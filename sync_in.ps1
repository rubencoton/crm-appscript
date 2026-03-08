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

Write-Host "[1/3] git pull --rebase"
try {
  & $git -C $repoRoot pull --rebase
  if ($LASTEXITCODE -ne 0) { throw "git pull --rebase devolvio codigo $LASTEXITCODE" }
} catch {
  Write-Host "ERROR de Git (sin forzar nada)."
  Write-Host "Elige: resolver conflictos manualmente, o abortar rebase con 'git rebase --abort'."
  throw
}

Write-Host "[2/3] clasp pull (Festivales)"
& $clasp -P $repoRoot pull
if ($LASTEXITCODE -ne 0) { throw "clasp pull en Festivales fallo" }

Write-Host "[3/3] clasp pull (CRM AYUDAS Y SUBVENCIONES)"
& $clasp -P $crmPath pull
if ($LASTEXITCODE -ne 0) { throw "clasp pull en CRM fallo" }

Write-Host "OK sync_in completado"
