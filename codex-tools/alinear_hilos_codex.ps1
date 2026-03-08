param(
  [string]$CanonicalWorkspace = "C:\Users\elrub\Desktop\CARPETA CODEX"
)

$ErrorActionPreference = "Stop"
$codexHome = Join-Path $env:USERPROFILE ".codex"
$statePath = Join-Path $codexHome ".codex-global-state.json"
if (-not (Test-Path $statePath)) { throw "No existe $statePath" }

$backupDir = Join-Path $codexHome "backups"
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupPath = Join-Path $backupDir (".codex-global-state_" + $stamp + ".json")
Copy-Item $statePath $backupPath -Force

$state = Get-Content $statePath -Raw | ConvertFrom-Json
if (-not $state.'electron-persisted-atom-state') {
  throw "Formato inesperado en .codex-global-state.json"
}

$existingActive = @($state.'active-workspace-roots')
$existingSaved = @($state.'electron-saved-workspace-roots')
$merged = @($CanonicalWorkspace) + $existingActive + $existingSaved
$uniqueRoots = @()
foreach ($root in $merged) {
  if ([string]::IsNullOrWhiteSpace($root)) { continue }
  if (-not ($uniqueRoots -contains $root)) {
    $uniqueRoots += $root
  }
}

$state.'active-workspace-roots' = $uniqueRoots
$state.'electron-saved-workspace-roots' = $uniqueRoots

$labels = $state.'electron-workspace-root-labels'
if ($null -eq $labels) {
  $labels = [pscustomobject]@{}
  if ($state.PSObject.Properties.Name -contains 'electron-workspace-root-labels') {
    $state.'electron-workspace-root-labels' = $labels
  } else {
    $state | Add-Member -NotePropertyName 'electron-workspace-root-labels' -NotePropertyValue $labels
  }
}

$labelNames = @($labels.PSObject.Properties | ForEach-Object { $_.Name })
if (-not ($labelNames -contains $CanonicalWorkspace)) {
  $labels | Add-Member -NotePropertyName $CanonicalWorkspace -NotePropertyValue "CARPETA CODEX" -Force
}

$state | ConvertTo-Json -Depth 100 -Compress | Set-Content -Path $statePath -Encoding UTF8

Write-Host "BACKUP_CREATED: $backupPath"
Write-Host "UPDATED_ACTIVE_ROOTS:"
$uniqueRoots | ForEach-Object { Write-Host "- $_" }
Write-Host "UPDATED_SAVED_ROOTS:"
$uniqueRoots | ForEach-Object { Write-Host "- $_" }
Write-Host "DONE"
