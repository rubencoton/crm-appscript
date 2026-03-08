param(
  [ValidateSet("login", "status", "pull", "push", "open", "version", "deploy")]
  [string]$Action = "status"
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

$nodeExe = Resolve-NodeExe

switch ($Action) {
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
