param(
  [string]$RepoUrl = "https://github.com/rubencoton/crm-appscript.git",
  [string]$BasePath = "C:\Users\$env:USERNAME\Desktop\CARPETA CODEX\01_PROYECTOS",
  [string]$RepoName = "festivales-github",
  [string]$ExpectedGoogleAccount = "",
  [switch]$SkipInstall,
  [switch]$SkipClaspLogin
)

$ErrorActionPreference = "Stop"
$script:GitExe = $null
$script:NodeExe = $null
$script:NpmExe = $null

function Write-Step {
  param([string]$Message)
  Write-Host "[SETUP] $Message"
}

function Resolve-Exe {
  param(
    [string]$Name,
    [string[]]$Candidates
  )

  $cmd = Get-Command $Name -ErrorAction SilentlyContinue
  if ($cmd -and $cmd.Source) {
    return $cmd.Source
  }

  foreach ($candidate in $Candidates) {
    if (Test-Path -LiteralPath $candidate) {
      return $candidate
    }
  }

  throw "No se encontro '$Name' ni en PATH ni en rutas candidatas."
}

try {
  $script:GitExe = Resolve-Exe -Name "git" -Candidates @(
    "C:\Program Files\Git\cmd\git.exe",
    "C:\Program Files (x86)\Git\cmd\git.exe"
  )
  $script:NodeExe = Resolve-Exe -Name "node" -Candidates @(
    "C:\Program Files\nodejs\node.exe",
    "C:\Program Files (x86)\nodejs\node.exe"
  )
  $script:NpmExe = Resolve-Exe -Name "npm" -Candidates @(
    "C:\Program Files\nodejs\npm.cmd",
    "C:\Program Files (x86)\nodejs\npm.cmd"
  )

  $env:Path = "$(Split-Path -Parent $script:GitExe);$(Split-Path -Parent $script:NodeExe);$env:Path"

  if (-not (Test-Path -LiteralPath $BasePath)) {
    Write-Step "Creando base: $BasePath"
    New-Item -ItemType Directory -Path $BasePath -Force | Out-Null
  }

  $repoPath = Join-Path $BasePath $RepoName
  if (-not (Test-Path -LiteralPath $repoPath)) {
    Write-Step "Clonando repo: $RepoUrl"
    & $script:GitExe clone $RepoUrl $repoPath
    if ($LASTEXITCODE -ne 0) {
      throw "Fallo git clone"
    }
  }
  else {
    Write-Step "Repo ya existe: $repoPath"
  }

  Push-Location $repoPath

  if (-not $SkipInstall) {
    Write-Step "Instalando dependencias npm"
    & $script:NpmExe install
    if ($LASTEXITCODE -ne 0) {
      throw "Fallo npm install"
    }
  }

  $claspLocal = Join-Path $repoPath "node_modules\@google\clasp\build\src\index.js"
  if (-not (Test-Path -LiteralPath $claspLocal)) {
    throw "No se encontro clasp local. Ejecuta npm install de nuevo."
  }

  Write-Step "Comprobando login de clasp"
  $authRaw = & $script:NodeExe $claspLocal show-authorized-user 2>&1
  $authText = ($authRaw | Out-String).Trim()
  $authUnknown = $authText -match "unknown user"
  $authFailed = $LASTEXITCODE -ne 0

  if ($authFailed -or $authUnknown) {
    if ($SkipClaspLogin) {
      throw "clasp no esta validado y SkipClaspLogin esta activo."
    }

    Write-Step "Iniciando login interactivo de clasp (primera vez por equipo)"
    & $script:NodeExe $claspLocal login
    if ($LASTEXITCODE -ne 0) {
      throw "Fallo clasp login"
    }

    $authRaw = & $script:NodeExe $claspLocal show-authorized-user 2>&1
    $authText = ($authRaw | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -or ($authText -match "unknown user")) {
      throw "Login completado pero la cuenta no queda validada."
    }
  }

  if (-not [string]::IsNullOrWhiteSpace($ExpectedGoogleAccount)) {
    if ($authText -notmatch [regex]::Escape($ExpectedGoogleAccount)) {
      throw "Cuenta clasp distinta a la esperada. Esperada: $ExpectedGoogleAccount. Detectada: $authText"
    }
  }

  Write-Step "Cuenta clasp detectada: $authText"
  Write-Step "Verificando proyecto clasp"
  & $script:NodeExe $claspLocal status
  if ($LASTEXITCODE -ne 0) {
    throw "Fallo clasp status"
  }

  Write-Step "Equipo listo en: $repoPath"
}
finally {
  Pop-Location -ErrorAction SilentlyContinue
}
