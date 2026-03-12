param(
  [ValidateSet("login", "status", "pull", "push", "open", "version", "deploy")]
  [string]$Action = "status",
  [switch]$ForceInteractiveLogin
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$claspLocal = Join-Path $repoRoot "node_modules\@google\clasp\build\src\index.js"

if (-not (Test-Path -LiteralPath $claspLocal)) {
  throw "No se encontro clasp local en $claspLocal"
}

function Resolve-NodeExe {
  $cmd = Get-Command node.exe -ErrorAction SilentlyContinue
  if ($cmd -and $cmd.Source) {
    return $cmd.Source
  }

  $candidates = @(
    "C:\Program Files\nodejs\node.exe",
    "C:\Program Files (x86)\nodejs\node.exe",
    "C:\Progra~1\nodejs\node.exe"
  )

  foreach ($candidate in $candidates) {
    if (Test-Path -LiteralPath $candidate) {
      return $candidate
    }
  }

  throw "No se encontro node.exe"
}

function Try-NonInteractiveLogin {
  $rcPath = Join-Path $env:USERPROFILE ".clasprc.json"
  if (-not (Test-Path -LiteralPath $rcPath)) {
    return $false
  }

  try {
    $cfg = Get-Content -LiteralPath $rcPath -Raw | ConvertFrom-Json
    $tok = $cfg.tokens.default
    if (-not $tok) { return $false }
    if (-not $tok.client_id -or -not $tok.client_secret -or -not $tok.refresh_token) { return $false }

    $refresh = Invoke-RestMethod -Method Post -Uri "https://oauth2.googleapis.com/token" -Body @{
      client_id = $tok.client_id
      client_secret = $tok.client_secret
      refresh_token = $tok.refresh_token
      grant_type = "refresh_token"
    }

    if (-not $refresh.access_token) {
      return $false
    }

    $headers = @{ Authorization = "Bearer $($refresh.access_token)" }
    $user = Invoke-RestMethod -Method Get -Uri "https://www.googleapis.com/oauth2/v2/userinfo" -Headers $headers
    $email = [string]$user.email
    if ([string]::IsNullOrWhiteSpace($email)) { $email = "(sin email)" }
    Write-Host "[INFO] Login no interactivo OK con: $email"
    return $true
  }
  catch {
    return $false
  }
}

$nodeExe = Resolve-NodeExe

switch ($Action) {
  "login" {
    if (-not $ForceInteractiveLogin) {
      if (Try-NonInteractiveLogin) { return }
      throw "No hay token reutilizable para login no interactivo. Por seguridad no se abre OAuth en navegador (evita bloqueo org_internal). Usa -ForceInteractiveLogin para forzar login manual."
    }

    & $nodeExe $claspLocal login
    if ($LASTEXITCODE -ne 0) { throw "clasp login fallo" }
  }
  "deploy" {
    & $nodeExe $claspLocal version
    if ($LASTEXITCODE -ne 0) { throw "clasp version fallo" }
    & $nodeExe $claspLocal deploy
    if ($LASTEXITCODE -ne 0) { throw "clasp deploy fallo" }
  }
  default {
    & $nodeExe $claspLocal $Action
    if ($LASTEXITCODE -ne 0) { throw "clasp $Action fallo" }
  }
}
