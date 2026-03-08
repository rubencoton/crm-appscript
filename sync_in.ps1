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

Write-Host "[1/3] git pull --ff-only"
try {
  & $git -C $repoRoot pull --ff-only origin main
  if ($LASTEXITCODE -ne 0) { throw "git pull --ff-only devolvio codigo $LASTEXITCODE" }
} catch {
  Write-Host "ERROR de Git (sin forzar nada)."
  Write-Host "Elige: resolver conflictos manualmente, o revisar estado con 'git status'."
  throw
}

Write-Host "[2/3] clasp pull (Festivales)"
$code = Invoke-Clasp -RepoRoot $repoRoot -ProjectPath $repoRoot -Action "pull"
if ($code -ne 0) { throw "clasp pull en Festivales fallo" }

Write-Host "[3/3] clasp pull (CRM AYUDAS Y SUBVENCIONES)"
$code = Invoke-Clasp -RepoRoot $repoRoot -ProjectPath $crmPath -Action "pull"
if ($code -ne 0) { throw "clasp pull en CRM fallo" }

Write-Host "OK sync_in completado"
