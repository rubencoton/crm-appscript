param(
  [Parameter(Mandatory = $true)]
  [string]$ScriptPath
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$nodePath = $null

try {
  $nodePath = (Get-Command node -ErrorAction Stop).Source
} catch {
  $candidates = @(
    "$env:ProgramFiles\nodejs\node.exe",
    "$env:ProgramFiles(x86)\nodejs\node.exe",
    "$env:LOCALAPPDATA\Programs\nodejs\node.exe"
  )

  foreach ($candidate in $candidates) {
    if ($candidate -and (Test-Path -LiteralPath $candidate)) {
      $nodePath = $candidate
      break
    }
  }
}

if (-not $nodePath) {
  throw 'No se encontro Node.js. Instala Node LTS o agrega node al PATH.'
}

$target = $ScriptPath
if (-not [System.IO.Path]::IsPathRooted($target)) {
  $target = Join-Path $repoRoot $target
}

$resolved = (Resolve-Path -LiteralPath $target).Path
& $nodePath $resolved
exit $LASTEXITCODE
