$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$syncIn = Join-Path $repoRoot "sync_in.ps1"
$syncOut = Join-Path $repoRoot "sync_out.ps1"

Write-Host "[1/2] Ejecutando sync_in.ps1"
& powershell -NoProfile -ExecutionPolicy Bypass -File $syncIn
if ($LASTEXITCODE -ne 0) { throw "sync_in fallo. Se detiene sync_full." }

Write-Host "[2/2] Ejecutando sync_out.ps1"
& powershell -NoProfile -ExecutionPolicy Bypass -File $syncOut
if ($LASTEXITCODE -ne 0) { throw "sync_out fallo." }

Write-Host "OK sync_full completado"
