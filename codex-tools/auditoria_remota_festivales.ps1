param(
  [string]$SpreadsheetId = "1kRrdCwd0n6FwVp-8rKP3gEeYT-TBXc1JL8sa_xzp1IM"
)

$ErrorActionPreference = "Stop"

function Resolve-NpmCmd {
  $cmd = Get-Command npm -ErrorAction SilentlyContinue
  if ($cmd -and $cmd.Source) { return $cmd.Source }
  $fallback = "C:\Program Files\nodejs\npm.cmd"
  if (Test-Path -LiteralPath $fallback) { return $fallback }
  throw "No se encontro npm.cmd"
}

function Run-Step {
  param(
    [string]$Name,
    [scriptblock]$Action
  )
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  try {
    $null = & $Action
    if ($LASTEXITCODE -ne $null -and $LASTEXITCODE -ne 0) {
      throw ("ExitCode=" + $LASTEXITCODE)
    }
    $sw.Stop()
    return [pscustomobject]@{
      name = $Name
      status = "OK"
      seconds = [math]::Round($sw.Elapsed.TotalSeconds, 2)
      error = ""
    }
  } catch {
    $sw.Stop()
    return [pscustomobject]@{
      name = $Name
      status = "ERROR"
      seconds = [math]::Round($sw.Elapsed.TotalSeconds, 2)
      error = $_.Exception.Message
    }
  }
}

$repo = Split-Path -Parent $PSScriptRoot
$reports = Join-Path $PSScriptRoot "reports"
if (-not (Test-Path -LiteralPath $reports)) { New-Item -ItemType Directory -Path $reports | Out-Null }

$npm = Resolve-NpmCmd
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$xlsx = Join-Path $reports ("sheet_snapshot_festivales_direct_" + $stamp + ".xlsx")

$steps = @()
$steps += Run-Step -Name "audit:auto" -Action { & $npm run audit:auto }
$steps += Run-Step -Name "audit:ultra" -Action { & $npm run audit:ultra }
$steps += Run-Step -Name "snapshot:xlsx" -Action {
  & curl.exe -L ("https://docs.google.com/spreadsheets/d/" + $SpreadsheetId + "/export?format=xlsx") -o $xlsx
}

$fallbackJson = ""
$fallbackMd = ""
$steps += Run-Step -Name "audit:sheet_fallback" -Action {
  $out = & python (Join-Path $PSScriptRoot "auditar_sheet_fallback.py") --spreadsheet-id $SpreadsheetId --xlsx-path $xlsx
  $obj = $out | ConvertFrom-Json
  $script:fallbackJson = $obj.json
  $script:fallbackMd = $obj.md
}

$steps += Run-Step -Name "gas:push" -Action { & $npm run gas:push }
$steps += Run-Step -Name "gas:deploy" -Action { & $npm run gas:deploy }

$summary = @{
  generatedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
  spreadsheetId = $SpreadsheetId
  xlsxPath = $xlsx
  fallbackJson = $fallbackJson
  fallbackMd = $fallbackMd
  steps = $steps
}

$jsonPath = Join-Path $reports ("auditoria_remota_festivales_" + $stamp + ".json")
$mdPath = Join-Path $reports ("auditoria_remota_festivales_" + $stamp + ".md")
$summary | ConvertTo-Json -Depth 10 | Set-Content -Encoding UTF8 $jsonPath

$lines = @(
  "# AUDITORIA REMOTA FESTIVALES",
  "",
  "- Fecha: " + $summary.generatedAt,
  "- Spreadsheet: " + $SpreadsheetId,
  "- Snapshot XLSX: " + $xlsx,
  "- Fallback JSON: " + $fallbackJson,
  "- Fallback MD: " + $fallbackMd,
  "",
  "## Pasos",
  ""
)
foreach ($s in $steps) {
  $line = "- " + $s.name + ": " + $s.status + " (" + $s.seconds + "s)"
  if ($s.error) { $line += " | " + $s.error }
  $lines += $line
}
$lines | Set-Content -Encoding UTF8 $mdPath

Write-Output ("JSON=" + $jsonPath)
Write-Output ("MD=" + $mdPath)
