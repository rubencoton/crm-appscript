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

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$crmPath = Join-Path $repoRoot "crm-ayudas-subvenciones"
$logPath = Join-Path $repoRoot "logs\sync_auto.log"
$lockPath = Join-Path $repoRoot "logs\sync_auto.lock"

function Write-Log([string]$msg) {
  $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg"
  Add-Content -Path $logPath -Value $line -Encoding UTF8
}

if (Test-Path $lockPath) {
  $ageMinutes = ((Get-Date) - (Get-Item $lockPath).LastWriteTime).TotalMinutes
  if ($ageMinutes -lt 20) {
    Write-Log "SKIP: ya hay ejecucion en curso."
    exit 0
  }
  Remove-Item $lockPath -Force
}

New-Item -ItemType File -Path $lockPath -Force | Out-Null

try {
  $git = Get-GitCmd
  $clasp = Get-ClaspCmd

  Set-Location $repoRoot

  $dirty = & $git -C $repoRoot status --porcelain
  if ($dirty) {
    Write-Log "SKIP: hay cambios locales. No se hace pull automatico para no interferir."
    exit 0
  }

  Write-Log "START: git pull --rebase"
  & $git -C $repoRoot pull --rebase
  if ($LASTEXITCODE -ne 0) {
    Write-Log "ERROR: git pull --rebase fallo con codigo $LASTEXITCODE"
    exit $LASTEXITCODE
  }

  Write-Log "STEP: clasp pull Festivales"
  & $clasp -P $repoRoot pull
  if ($LASTEXITCODE -ne 0) {
    Write-Log "ERROR: clasp pull Festivales fallo con codigo $LASTEXITCODE"
    exit $LASTEXITCODE
  }

  Write-Log "STEP: clasp pull CRM"
  & $clasp -P $crmPath pull
  if ($LASTEXITCODE -ne 0) {
    Write-Log "ERROR: clasp pull CRM fallo con codigo $LASTEXITCODE"
    exit $LASTEXITCODE
  }

  Write-Log "OK: sync_in_auto completado"
  exit 0
} catch {
  Write-Log ("FATAL: " + $_.Exception.Message)
  exit 1
} finally {
  if (Test-Path $lockPath) {
    Remove-Item $lockPath -Force
  }
}
