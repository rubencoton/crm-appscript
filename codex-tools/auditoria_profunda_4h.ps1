param(
  [int]$Horas = 4,
  [int]$IntervaloMin = 10,
  [string]$SpreadsheetId = '1LgZG2ObSCJzEQvrysu1NFFEvYlupLXVByDnIMCr-wYA',
  [string]$ProjectDir = 'C:\Users\elrub\Desktop\CARPETA CODEX\01_PROYECTOS\festivales-github\crm-ayudas-subvenciones',
  [string]$Clasprc = 'C:\Users\elrub\Desktop\CARPETA CODEX\01_PROYECTOS\noches-neon-crm-gas\.clasprc.json'
)

$ErrorActionPreference = 'Continue'

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$reportsDir = Join-Path $baseDir 'reports'
if (-not (Test-Path -LiteralPath $reportsDir)) {
  New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null
}

$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = Join-Path $reportsDir ("auditoria_4h_log_{0}.txt" -f $runId)
$csvPath = Join-Path $reportsDir ("auditoria_4h_resumen_{0}.csv" -f $runId)

"timestamp,findings,critical,high,medium,low,audit_json,audit_md" | Set-Content -LiteralPath $csvPath -Encoding UTF8

function Write-RunLog {
  param([string]$Message)
  $line = ("[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
  Add-Content -LiteralPath $logPath -Value $line
}

Write-RunLog "INICIO auditoria profunda. Horas=$Horas IntervaloMin=$IntervaloMin SpreadsheetId=$SpreadsheetId"

$endAt = (Get-Date).AddHours($Horas)
$cycle = 0

while ((Get-Date) -lt $endAt) {
  $cycle++
  Write-RunLog ("CICLO {0} - inicio" -f $cycle)

  $cmd = @(
    'python',
    'auditar_crm_ayudas_extremo.py',
    '--spreadsheet-id', $SpreadsheetId,
    '--project-dir', $ProjectDir,
    '--clasprc', $Clasprc
  )

  Push-Location $baseDir
  $output = & $cmd[0] $cmd[1] $cmd[2] $cmd[3] $cmd[4] $cmd[5] $cmd[6] $cmd[7] 2>&1
  $exit = $LASTEXITCODE
  Pop-Location

  if ($exit -ne 0) {
    Write-RunLog ("CICLO {0} - ERROR auditoria (exit={1}) :: {2}" -f $cycle, $exit, ($output -join ' | '))
  } else {
    $raw = ($output -join '')
    try {
      $obj = $raw | ConvertFrom-Json
      $jsonPath = [string]$obj.json
      $mdPath = [string]$obj.md
      $findings = [int]$obj.findings
      $critical = 0
      $high = 0
      $medium = 0
      $low = 0

      if (Test-Path -LiteralPath $jsonPath) {
        $audit = Get-Content -LiteralPath $jsonPath -Raw | ConvertFrom-Json
        foreach ($f in @($audit.findings)) {
          $sev = [string]$f.severity
          if ($sev -eq 'CRITICAL') { $critical++ }
          elseif ($sev -eq 'HIGH') { $high++ }
          elseif ($sev -eq 'MEDIUM') { $medium++ }
          elseif ($sev -eq 'LOW') { $low++ }
        }
      }

      $row = ('"{0}",{1},{2},{3},{4},{5},"{6}","{7}"' -f
        (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
        $findings, $critical, $high, $medium, $low,
        ($jsonPath -replace '"', '""'),
        ($mdPath -replace '"', '""'))
      Add-Content -LiteralPath $csvPath -Value $row
      Write-RunLog ("CICLO {0} - OK findings={1} (C={2} H={3} M={4} L={5})" -f $cycle, $findings, $critical, $high, $medium, $low)
    }
    catch {
      Write-RunLog ("CICLO {0} - ERROR parse output :: {1}" -f $cycle, $_.Exception.Message)
      Write-RunLog ("CICLO {0} - RAW :: {1}" -f $cycle, $raw)
    }
  }

  $now = Get-Date
  if ($now -ge $endAt) { break }
  $sleepSec = [Math]::Max(15, $IntervaloMin * 60)
  Write-RunLog ("CICLO {0} - espera {1} segundos" -f $cycle, $sleepSec)
  Start-Sleep -Seconds $sleepSec
}

Write-RunLog "FIN auditoria profunda."
Write-Output ("LOG={0}" -f $logPath)
Write-Output ("CSV={0}" -f $csvPath)
