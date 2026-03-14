param(
  [int]$Horas = 2,
  [int]$IntervaloMin = 10
)

$ErrorActionPreference = "Continue"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoDir = Split-Path -Parent $baseDir
$reportsDir = Join-Path $baseDir "reports"
if (-not (Test-Path -LiteralPath $reportsDir)) {
  New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null
}

$runId = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath = Join-Path $reportsDir ("auditoria_festivales_2h_log_{0}.txt" -f $runId)
$csvPath = Join-Path $reportsDir ("auditoria_festivales_2h_resumen_{0}.csv" -f $runId)
"timestamp,cycle,type,exit_code,pass,fail,notes" | Set-Content -LiteralPath $csvPath -Encoding UTF8

function Write-RunLog {
  param([string]$Message)
  $line = ("[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message)
  Add-Content -LiteralPath $logPath -Value $line
}

function Resolve-Node {
  $cmd = Get-Command node.exe -ErrorAction SilentlyContinue
  if ($cmd -and $cmd.Source) { return $cmd.Source }
  $fallback = "C:\Program Files\nodejs\node.exe"
  if (Test-Path -LiteralPath $fallback) { return $fallback }
  throw "No se encontro node.exe"
}

function Parse-AuditResult {
  param([string]$Text)
  $pass = 0
  $fail = 0
  if ($Text -match "Pass:\s*([0-9]+)") { $pass = [int]$Matches[1] }
  if ($Text -match "Fail:\s*([0-9]+)") { $fail = [int]$Matches[1] }
  return @{ pass = $pass; fail = $fail }
}

$nodeExe = Resolve-Node
$auditAuto = Join-Path $baseDir "audit_auto.js"
$auditUltra = Join-Path $baseDir "audit_ultra.js"

Write-RunLog ("INICIO auditoria festivales 2h. IntervaloMin={0}" -f $IntervaloMin)

$endAt = (Get-Date).AddHours($Horas)
$cycle = 0

while ((Get-Date) -lt $endAt) {
  $cycle++
  Write-RunLog ("CICLO {0} inicio" -f $cycle)

  Push-Location $repoDir

  $autoOut = & $nodeExe $auditAuto 2>&1
  $autoExit = $LASTEXITCODE
  $autoTxt = ($autoOut -join [Environment]::NewLine)
  $autoRes = Parse-AuditResult -Text $autoTxt
  $autoNotes = ("audit_auto pass={0} fail={1}" -f $autoRes.pass, $autoRes.fail)
  Add-Content -LiteralPath $csvPath -Value ('"{0}",{1},"audit_auto",{2},{3},{4},"{5}"' -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $cycle, $autoExit, $autoRes.pass, $autoRes.fail, ($autoNotes -replace '"','""'))
  Write-RunLog ("CICLO {0} audit_auto exit={1} pass={2} fail={3}" -f $cycle, $autoExit, $autoRes.pass, $autoRes.fail)

  $ultraOut = & $nodeExe $auditUltra 2>&1
  $ultraExit = $LASTEXITCODE
  $ultraTxt = ($ultraOut -join [Environment]::NewLine)
  $ultraRes = Parse-AuditResult -Text $ultraTxt
  $ultraNotes = ("audit_ultra pass={0} fail={1}" -f $ultraRes.pass, $ultraRes.fail)
  Add-Content -LiteralPath $csvPath -Value ('"{0}",{1},"audit_ultra",{2},{3},{4},"{5}"' -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $cycle, $ultraExit, $ultraRes.pass, $ultraRes.fail, ($ultraNotes -replace '"','""'))
  Write-RunLog ("CICLO {0} audit_ultra exit={1} pass={2} fail={3}" -f $cycle, $ultraExit, $ultraRes.pass, $ultraRes.fail)

  $gitStatus = & git status --short
  if ($LASTEXITCODE -eq 0 -and ($gitStatus -join "")) {
    Write-RunLog ("CICLO {0} cambios detectados por auditoria, creando commit de trazabilidad." -f $cycle)
    & git add -A
    & git commit -m ("Festivales: trazabilidad auditoria 2h ciclo {0}" -f $cycle) 2>&1 | Out-Null
    & git push origin main 2>&1 | Out-Null
  }

  Pop-Location

  if ((Get-Date) -ge $endAt) { break }
  $sleepSec = [Math]::Max(60, $IntervaloMin * 60)
  Write-RunLog ("CICLO {0} espera {1} segundos" -f $cycle, $sleepSec)
  Start-Sleep -Seconds $sleepSec
}

Write-RunLog "FIN auditoria festivales 2h."
Write-Output ("LOG={0}" -f $logPath)
Write-Output ("CSV={0}" -f $csvPath)
