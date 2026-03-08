param(
  [switch]$DryRun,
  [switch]$AllowDirty,
  [switch]$SkipCrmProject
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

if ((-not $DryRun) -and (-not $AllowDirty)) {
  $dirty = & $git -c "safe.directory=$repoRoot" -C $repoRoot status --porcelain
  if (-not [string]::IsNullOrWhiteSpace(($dirty | Out-String))) {
    throw "Hay cambios locales. Usa -AllowDirty o limpia el arbol antes de sync_in."
  }
}

$code = Run-OrDry -Title "[1/3] git pull --ff-only" -Action {
  & $git -c "safe.directory=$repoRoot" -C $repoRoot pull --ff-only origin main
}
if ($code -ne 0) {
  Write-Host "ERROR de Git (sin forzar nada)."
  Write-Host "Elige: resolver conflictos manualmente, o revisar estado con 'git status'."
  throw "git pull --ff-only devolvio codigo $code"
}

$code = Run-OrDry -Title "[2/3] clasp pull (Festivales)" -Action {
  Invoke-Clasp -RepoRoot $repoRoot -ProjectPath $repoRoot -Action "pull"
}
if ($code -ne 0) { throw "clasp pull en Festivales fallo" }

if ($SkipCrmProject) {
  Write-Host "[3/3] omitido por -SkipCrmProject"
} else {
  if (-not (Test-Path -LiteralPath (Join-Path $crmPath ".clasp.json"))) {
    throw "No existe .clasp.json en proyecto CRM secundario: $crmPath"
  }
  $code = Run-OrDry -Title "[3/3] clasp pull (CRM AYUDAS Y SUBVENCIONES)" -Action {
    Invoke-Clasp -RepoRoot $repoRoot -ProjectPath $crmPath -Action "pull"
  }
  if ($code -ne 0) { throw "clasp pull en CRM fallo" }
}

Write-Host "OK sync_in completado"
