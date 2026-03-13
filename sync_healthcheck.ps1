$ErrorActionPreference = "Stop"

function Out-Line([string]$label, [string]$value) {
  "$label`t$value"
}

function Get-WinGetNodeDir {
  $packagesRoot = Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Packages"
  if (-not (Test-Path $packagesRoot)) { return $null }

  $pkg = Get-ChildItem -Path $packagesRoot -Directory -Filter "OpenJS.NodeJS.LTS*" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
  if (-not $pkg) { return $null }

  $nodeDir = Get-ChildItem -Path $pkg.FullName -Directory -Filter "node-v*-win-x64" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

  if ($nodeDir) { return $nodeDir.FullName }
  return $null
}

function Ensure-ToolPath {
  $wingetNode = Get-WinGetNodeDir
  $candidates = @(
    "C:\Program Files\nodejs",
    "C:\Program Files\Git\cmd",
    "C:\Program Files\Git\mingw64\bin",
    "$env:LOCALAPPDATA\Microsoft\WinGet\Links",
    $wingetNode
  )

  foreach ($dir in $candidates) {
    if ([string]::IsNullOrWhiteSpace($dir)) { continue }
    if (-not (Test-Path $dir)) { continue }
    if ($env:Path -notlike "*$dir*") {
      $env:Path = "$dir;$env:Path"
    }
  }
}

function Resolve-ToolPath {
  param(
    [string]$Executable,
    [string[]]$Fallbacks = @()
  )

  $cmd = Get-Command $Executable -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }

  foreach ($path in $Fallbacks) {
    if ([string]::IsNullOrWhiteSpace($path)) { continue }
    if (Test-Path $path) { return $path }
  }

  return $null
}

function Test-CommandExecutable {
  param(
    [string]$CommandPath,
    [string[]]$CommandArgs = @("--version")
  )

  if ([string]::IsNullOrWhiteSpace($CommandPath)) { return $false }
  try {
    & $CommandPath @CommandArgs *> $null
    return ($LASTEXITCODE -eq 0)
  } catch {
    return $false
  }
}

function Invoke-OrMissing {
  param(
    [string]$CommandPath,
    [string[]]$Arguments = @()
  )

  if ([string]::IsNullOrWhiteSpace($CommandPath)) {
    return "MISSING"
  }

  try {
    return (& $CommandPath @Arguments 2>$null)
  } catch {
    return "ERROR"
  }
}

function Resolve-ClaspCmd {
  param(
    [string]$RepoRoot,
    [string]$WinGetNodeDir
  )

  $localClasp = Join-Path $RepoRoot "node_modules\.bin\clasp.cmd"
  $roamingClasp = Join-Path $env:APPDATA "npm\clasp.cmd"

  $winGetClasp = $null
  if ($WinGetNodeDir) {
    $winGetClasp = Join-Path $WinGetNodeDir "clasp.cmd"
  }
  $fromPath = Resolve-ToolPath -Executable "clasp.cmd"
  $fallbacks = @(
    $fromPath,
    $roamingClasp,
    $localClasp,
    (Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Packages\OpenJS.NodeJS.LTS_Microsoft.Winget.Source_8wekyb3d8bbwe\node-v24.14.0-win-x64\clasp.cmd"),
    $winGetClasp
  ) | Select-Object -Unique

  foreach ($candidate in $fallbacks) {
    if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
    if (Test-Path $candidate) {
      return $candidate
    }
  }

  return $null
}

function Resolve-NpxCmd {
  param([string]$WinGetNodeDir)

  $winGetNpx = $null
  if ($WinGetNodeDir) {
    $winGetNpx = Join-Path $WinGetNodeDir "npx.cmd"
  }
  $fallbacks = @(
    "C:\Program Files\nodejs\npx.cmd",
    $winGetNpx
  )

  return Resolve-ToolPath -Executable "npx.cmd" -Fallbacks $fallbacks
}

function Invoke-ClaspInfo {
  param(
    [string]$ClaspCmd,
    [string]$NpxCmd,
    [string[]]$CommandArgs
  )

  if (-not [string]::IsNullOrWhiteSpace($ClaspCmd)) {
    try {
      $out = (& $ClaspCmd @CommandArgs 2>$null)
      return $out
    } catch {
      # fallback to npx
    }
  }

  if (-not [string]::IsNullOrWhiteSpace($NpxCmd)) {
    try {
      $out = (& $NpxCmd "clasp" @CommandArgs 2>$null)
      return $out
    } catch {
      # continue to second attempt
    }

    try {
      $out = (& $NpxCmd "--yes" "@google/clasp" @CommandArgs 2>$null)
      return $out
    } catch {
      # no-op
    }
  }

  return "ERROR"
}

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$reportDir = Join-Path $repoRoot "logs"
$reportPath = Join-Path $reportDir "healthcheck-latest.txt"
New-Item -ItemType Directory -Path $reportDir -Force | Out-Null

$expectedSheetId = "1kRrdCwd0n6FwVp-8rKP3gEeYT-TBXc1JL8sa_xzp1IM"
$expectedScriptId = "1OGuPezQ26BFvaLRiy-IYIotGpmVu_Z_b9Mi8tCiprIz8zB4DgqmMc5Ea"
$expectedOrigin = "https://github.com/rubencoton/crm-appscript.git"
$expectedTask = "\CodexSyncIn_5min"
$legacyTask = "\CodexSyncInFestivales_Every5Min"

Ensure-ToolPath
$wingetNode = Get-WinGetNodeDir
$nodeFallback = $null
$npmFallback = $null
if ($wingetNode) {
  $nodeFallback = Join-Path $wingetNode "node.exe"
  $npmFallback = Join-Path $wingetNode "npm.cmd"
}
$git = Resolve-ToolPath -Executable "git.exe" -Fallbacks @("C:\Program Files\Git\cmd\git.exe")
$node = Resolve-ToolPath -Executable "node.exe" -Fallbacks @($nodeFallback)
$npm = Resolve-ToolPath -Executable "npm.cmd" -Fallbacks @("C:\Program Files\nodejs\npm.cmd", $npmFallback)
$clasp = Resolve-ClaspCmd -RepoRoot $repoRoot -WinGetNodeDir $wingetNode
$npx = Resolve-NpxCmd -WinGetNodeDir $wingetNode

$lines = @()
$lines += "HEALTHCHECK $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$lines += ""

$gitVer = Invoke-OrMissing -CommandPath $git -Arguments @("--version")
$nodeVer = Invoke-OrMissing -CommandPath $node -Arguments @("--version")
$npmVer = Invoke-OrMissing -CommandPath $npm -Arguments @("--version")
$claspVer = Invoke-ClaspInfo -ClaspCmd $clasp -NpxCmd $npx -CommandArgs @("--version")

$lines += Out-Line "git" "$gitVer"
$lines += Out-Line "node" "$nodeVer"
$lines += Out-Line "npm" "$npmVer"
$lines += Out-Line "clasp" "$claspVer"
$lines += Out-Line "clasp_cmd" "$clasp"
$lines += Out-Line "npx_cmd" "$npx"

$branch = "UNKNOWN"
$remote = "UNKNOWN"
$status = "UNKNOWN"
$last = "UNKNOWN"
if (-not [string]::IsNullOrWhiteSpace($git)) {
  try {
    $branch = (& $git -C $repoRoot rev-parse --abbrev-ref HEAD 2>$null)
    $remote = (& $git -C $repoRoot remote get-url origin 2>$null)
    $status = (& $git -C $repoRoot status --short --branch 2>$null)
    $last = (& $git -C $repoRoot log -1 --oneline 2>$null)
  } catch {
    # no-op
  }
}

$lines += ""
$lines += Out-Line "branch" "$branch"
$lines += Out-Line "remote" "$remote"
$lines += Out-Line "remote_expected" "$expectedOrigin"
$lines += Out-Line "remote_match" ([string]($remote -eq $expectedOrigin))
$lines += Out-Line "last_commit" "$last"
$lines += "git_status"
$lines += $status

$currentScriptId = "MISSING"
$claspJsonPath = Join-Path $repoRoot ".clasp.json"
if (Test-Path $claspJsonPath) {
  try {
    $claspJson = Get-Content $claspJsonPath -Raw | ConvertFrom-Json
    $currentScriptId = [string]$claspJson.scriptId
  } catch {
    $currentScriptId = "ERROR"
  }
}

$authUser = Invoke-ClaspInfo -ClaspCmd $clasp -NpxCmd $npx -CommandArgs @("show-authorized-user")
$lines += ""
$lines += Out-Line "clasp_user" (($authUser -join ' ').Trim())
$lines += Out-Line "scope_expected_sheet" "$expectedSheetId"
$lines += Out-Line "scope_expected_script" "$expectedScriptId"
$lines += Out-Line "scope_current_script" "$currentScriptId"
$lines += Out-Line "scope_script_match" ([string]($currentScriptId -eq $expectedScriptId))

$codexStatePath = Join-Path $env:USERPROFILE ".codex\.codex-global-state.json"
if (Test-Path $codexStatePath) {
  $st = Get-Content $codexStatePath -Raw | ConvertFrom-Json
  $cloud = $st.'electron-persisted-atom-state'.codexCloudAccess
  $activeRoots = @($st.'active-workspace-roots') -join ' | '
  $lines += Out-Line "codex_cloud" "$cloud"
  $lines += Out-Line "codex_active_roots" "$activeRoots"
}

$taskRaw = schtasks /Query /TN $expectedTask /FO LIST /V 2>$null
if (-not $taskRaw) {
  $taskRaw = schtasks /Query /TN $legacyTask /FO LIST /V 2>$null
}
$nextRun = ($taskRaw | Select-String -Pattern '(?i)next run|hora .*ejec').Line
$taskState = ($taskRaw | Select-String -Pattern '(?i)estado de tarea programada|scheduled task state|status').Line
$lastResult = ($taskRaw | Select-String -Pattern '(?i)ultimo resultado|último resultado|last result|resultado').Line
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
