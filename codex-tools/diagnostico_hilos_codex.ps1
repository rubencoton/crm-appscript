param(
  [string]$CanonicalWorkspace = "C:\Users\elrub\Desktop\CARPETA CODEX"
)

$ErrorActionPreference = "Stop"
$codexHome = Join-Path $env:USERPROFILE ".codex"
$authPath = Join-Path $codexHome "auth.json"
$statePath = Join-Path $codexHome ".codex-global-state.json"
$indexPath = Join-Path $codexHome "session_index.jsonl"
$sessionsDir = Join-Path $codexHome "sessions"

if (-not (Test-Path $authPath)) { throw "No existe $authPath" }
if (-not (Test-Path $statePath)) { throw "No existe $statePath" }

$auth = Get-Content $authPath -Raw | ConvertFrom-Json
$state = Get-Content $statePath -Raw | ConvertFrom-Json
$atom = $state.'electron-persisted-atom-state'

$cloud = $atom.codexCloudAccess
$activeRoots = @($state.'active-workspace-roots')
$savedRoots = @($state.'electron-saved-workspace-roots')
$sessionsCount = 0
$indexCount = 0

if (Test-Path $sessionsDir) {
  $sessionsCount = (Get-ChildItem $sessionsDir -Recurse -File | Measure-Object).Count
}
if (Test-Path $indexPath) {
  $indexCount = (Get-Content $indexPath | Measure-Object).Count
}

$canonicalInActive = $activeRoots -contains $CanonicalWorkspace
$canonicalInSaved = $savedRoots -contains $CanonicalWorkspace

Write-Host "ACCOUNT_ID: $($auth.tokens.account_id)"
Write-Host "AUTH_MODE: $($auth.auth_mode)"
Write-Host "CLOUD_ACCESS: $cloud"
Write-Host "CANONICAL_WORKSPACE: $CanonicalWorkspace"
Write-Host "CANONICAL_IN_ACTIVE: $canonicalInActive"
Write-Host "CANONICAL_IN_SAVED: $canonicalInSaved"
Write-Host "ACTIVE_ROOTS:"
$activeRoots | ForEach-Object { Write-Host "- $_" }
Write-Host "SAVED_ROOTS:"
$savedRoots | ForEach-Object { Write-Host "- $_" }
Write-Host "SESSION_INDEX_LINES: $indexCount"
Write-Host "SESSION_FILES: $sessionsCount"

if ($cloud -ne 'enabled') {
  Write-Host "NOTE: cloud sync de hilos no esta plenamente activado (valor actual: $cloud)."
}
