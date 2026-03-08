param(
  [switch]$DryRun,
  [switch]$AllowDirty,
  [switch]$SkipCrmProject,
  [switch]$SkipGitPush,
  [string]$CommitMessage = ""
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$syncIn = Join-Path $repoRoot "sync_in.ps1"
$syncOut = Join-Path $repoRoot "sync_out.ps1"

$inArgs = @()
if ($DryRun) { $inArgs += "-DryRun" }
if ($AllowDirty) { $inArgs += "-AllowDirty" }
if ($SkipCrmProject) { $inArgs += "-SkipCrmProject" }

$outArgs = @()
if ($DryRun) { $outArgs += "-DryRun" }
if ($SkipCrmProject) { $outArgs += "-SkipCrmProject" }
if ($SkipGitPush) { $outArgs += "-SkipGitPush" }
if (-not [string]::IsNullOrWhiteSpace($CommitMessage)) {
  $outArgs += "-CommitMessage"
  $outArgs += $CommitMessage
}

Write-Host "[1/2] Ejecutando sync_in.ps1"
& powershell -NoProfile -ExecutionPolicy Bypass -File $syncIn @inArgs
if ($LASTEXITCODE -ne 0) { throw "sync_in fallo. Se detiene sync_full." }

Write-Host "[2/2] Ejecutando sync_out.ps1"
& powershell -NoProfile -ExecutionPolicy Bypass -File $syncOut @outArgs
if ($LASTEXITCODE -ne 0) { throw "sync_out fallo." }

Write-Host "OK sync_full completado"
