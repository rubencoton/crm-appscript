param(
  [switch]$DryRun,
  [switch]$AllowDirty,
  [switch]$SkipCrmProject
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

if ((-not $DryRun) -and (-not $AllowDirty)) {
  $dirty = & $git -c "safe.directory=$repoRoot" -C $repoRoot status --porcelain
  if (-not [string]::IsNullOrWhiteSpace(($dirty | Out-String))) {
    throw "Hay cambios locales. Usa -AllowDirty o limpia el arbol antes de sync_in."
  }
}

try {
  Invoke-Step -Title "[1/3] git pull --ff-only" -ErrorMessage "git pull --ff-only fallo" -Action {
    & $git -c "safe.directory=$repoRoot" -C $repoRoot pull --ff-only origin main
  }
} catch {
  Write-Host "ERROR de Git (sin forzar nada)."
  Write-Host "Elige: resolver conflictos manualmente, o revisar estado con 'git status'."
  throw
}

Invoke-Step -Title "[2/3] clasp pull (Festivales)" -ErrorMessage "clasp pull en Festivales fallo" -Action {
  Invoke-Clasp -RepoRoot $repoRoot -ProjectPath $repoRoot -Action "pull"
}

if ($SkipCrmProject) {
  Write-Host "[3/3] omitido por -SkipCrmProject"
} else {
  if (-not (Test-Path -LiteralPath (Join-Path $crmPath ".clasp.json"))) {
    throw "No existe .clasp.json en proyecto CRM secundario: $crmPath"
  }

  Invoke-Step -Title "[3/3] clasp pull (CRM AYUDAS Y SUBVENCIONES)" -ErrorMessage "clasp pull en CRM fallo" -Action {
    Invoke-Clasp -RepoRoot $repoRoot -ProjectPath $crmPath -Action "pull"
  }
}

Write-Host "OK sync_in completado"
