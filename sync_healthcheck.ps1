$ErrorActionPreference = "Stop"

function Out-Line([string]$label, [string]$value) {
  "$label`t$value"
}

function Get-OrMissing([string]$path, [scriptblock]$cmd) {
  if (Test-Path $path) {
    return (& $cmd)
  }
  return "MISSING"
}

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$reportDir = Join-Path $repoRoot "logs"
$reportPath = Join-Path $reportDir "healthcheck-latest.txt"
New-Item -ItemType Directory -Path $reportDir -Force | Out-Null

$git = "C:\Program Files\Git\cmd\git.exe"
$node = "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\OpenJS.NodeJS.LTS_Microsoft.Winget.Source_8wekyb3d8bbwe\node-v24.14.0-win-x64\node.exe"
$npm = "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\OpenJS.NodeJS.LTS_Microsoft.Winget.Source_8wekyb3d8bbwe\node-v24.14.0-win-x64\npm.cmd"
$clasp = "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\OpenJS.NodeJS.LTS_Microsoft.Winget.Source_8wekyb3d8bbwe\node-v24.14.0-win-x64\clasp.cmd"

$lines = @()
$lines += "HEALTHCHECK $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$lines += ""

$gitVer = Get-OrMissing $git { & $git --version }
$nodeVer = Get-OrMissing $node { & $node --version }
$npmVer = Get-OrMissing $npm { & $npm --version }
$claspVer = Get-OrMissing $clasp { & $clasp --version }

$lines += Out-Line "git" "$gitVer"
$lines += Out-Line "node" "$nodeVer"
$lines += Out-Line "npm" "$npmVer"
$lines += Out-Line "clasp" "$claspVer"

$branch = & $git -C $repoRoot rev-parse --abbrev-ref HEAD
$remote = & $git -C $repoRoot remote get-url origin
$status = & $git -C $repoRoot status --short --branch
$last = & $git -C $repoRoot log -1 --oneline

$lines += ""
$lines += Out-Line "branch" "$branch"
$lines += Out-Line "remote" "$remote"
$lines += Out-Line "last_commit" "$last"
$lines += "git_status"
$lines += $status

$authUser = & $clasp show-authorized-user 2>$null
$lines += ""
$lines += Out-Line "clasp_user" (($authUser -join ' ').Trim())

$codexStatePath = Join-Path $env:USERPROFILE ".codex\.codex-global-state.json"
if (Test-Path $codexStatePath) {
  $st = Get-Content $codexStatePath -Raw | ConvertFrom-Json
  $cloud = $st.'electron-persisted-atom-state'.codexCloudAccess
  $activeRoots = @($st.'active-workspace-roots') -join ' | '
  $lines += Out-Line "codex_cloud" "$cloud"
  $lines += Out-Line "codex_active_roots" "$activeRoots"
}

$taskRaw = schtasks /Query /TN CodexSyncInFestivales_Every5Min /FO LIST /V 2>$null
$nextRun = ($taskRaw | Select-String -Pattern 'Hora próxima ejecución').Line
$taskState = ($taskRaw | Select-String -Pattern 'Estado de tarea programada').Line
$lastResult = ($taskRaw | Select-String -Pattern 'Último resultado').Line
$lines += Out-Line "task_every5_next" (($nextRun -replace '^.*?:\s*','').Trim())
$lines += Out-Line "task_every5_state" (($taskState -replace '^.*?:\s*','').Trim())
$lines += Out-Line "task_every5_last" (($lastResult -replace '^.*?:\s*','').Trim())

$startup = 'C:\Users\elrub\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\CodexSyncInOnLogon.bat'
$startupStatus = "MISSING"
if (Test-Path $startup) {
  $startupStatus = "OK"
}
$lines += Out-Line "startup_fallback" $startupStatus

Set-Content -Path $reportPath -Value $lines -Encoding UTF8
$lines | ForEach-Object { Write-Host $_ }
