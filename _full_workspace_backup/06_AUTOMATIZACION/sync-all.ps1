param(
    [ValidateSet("backup","full")]
    [string]$Mode = "backup",
    [string]$WorkspaceRoot = "C:\Users\elrub\Desktop\CARPETA CODEX",
    [string]$ProjectsRoot = ""
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ProjectsRoot)) {
    $ProjectsRoot = Join-Path $WorkspaceRoot "01_PROYECTOS"
}

# Ensure Node.js is available for clasp under scheduled execution.
$nodeDir = Join-Path $env:ProgramFiles "nodejs"
if (Test-Path $nodeDir) {
    $env:Path = $nodeDir + ";" + $env:Path
}

$autoDir = Join-Path $WorkspaceRoot "06_AUTOMATIZACION"
$logDir = Join-Path $autoDir "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$logFile = Join-Path $logDir ("sync-" + $timestamp + ".log")
$lockFile = Join-Path $logDir "sync.lock"

$gitCmd = (Get-Command git -ErrorAction SilentlyContinue).Source
if (-not $gitCmd -and (Test-Path "$env:ProgramFiles\Git\cmd\git.exe")) {
    $gitCmd = "$env:ProgramFiles\Git\cmd\git.exe"
}
if (-not $gitCmd -and (Test-Path "$env:ProgramFiles\Git\bin\git.exe")) {
    $gitCmd = "$env:ProgramFiles\Git\bin\git.exe"
}

$claspCmd = Join-Path $env:APPDATA "npm\clasp.cmd"
if (-not (Test-Path $claspCmd)) {
    $detectedClasp = (Get-Command clasp.cmd -ErrorAction SilentlyContinue).Source
    if ($detectedClasp) { $claspCmd = $detectedClasp }
}

if (Test-Path $lockFile) {
    $ageMinutes = ((Get-Date) - (Get-Item $lockFile).LastWriteTime).TotalMinutes
    if ($ageMinutes -gt 120) {
        Write-Host "[WARN] Removing stale lock file (>120 min)."
        Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue
    }
    else {
        Write-Host "[INFO] Another sync process is running. Exiting."
        exit 0
    }
}

New-Item -ItemType File -Path $lockFile -Force | Out-Null
Start-Transcript -Path $logFile -Append | Out-Null

function Get-GitRepos {
    param([string]$BasePath)

    $repos = @()

    if (Test-Path (Join-Path $BasePath ".git")) {
        $repos += $BasePath
    }

    if (Test-Path $BasePath) {
        $repos += Get-ChildItem -Path $BasePath -Directory -Force -Recurse -Filter ".git" -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Parent.FullName }
    }

    return $repos | Sort-Object -Unique
}

function Invoke-Git {
    param(
        [string]$RepoPath,
        [string[]]$GitArgs,
        [int[]]$AllowedExitCodes = @(0)
    )

    if (-not $script:gitCmd) {
        throw "git executable not found."
    }

    $oldEap = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $output = & $script:gitCmd -C $RepoPath @GitArgs 2>&1
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $oldEap
    }

    if ($AllowedExitCodes -contains $exitCode) {
        return [PSCustomObject]@{
            ExitCode = $exitCode
            Output   = $output
        }
    }

    $joined = ($output | Out-String).Trim()
    if ($joined -match "detected dubious ownership") {
        Write-Host "[WARN] Adding safe.directory for repo and retrying once: $RepoPath"
        & $script:gitCmd config --global --add safe.directory $RepoPath 2>&1 | Out-Null

        $oldEap = $ErrorActionPreference
        try {
            $ErrorActionPreference = "Continue"
            $output = & $script:gitCmd -C $RepoPath @GitArgs 2>&1
            $exitCode = $LASTEXITCODE
        }
        finally {
            $ErrorActionPreference = $oldEap
        }

        if ($AllowedExitCodes -contains $exitCode) {
            return [PSCustomObject]@{
                ExitCode = $exitCode
                Output   = $output
            }
        }
        $joined = ($output | Out-String).Trim()
    }

    throw "git failed ($($GitArgs -join ' ')) in $RepoPath. ExitCode=$exitCode. $joined"
}

function Get-GitText {
    param([object]$GitResult)
    return ($GitResult.Output | Out-String).Trim()
}

function Ensure-LocalExclude {
    param([string]$RepoPath)

    $excludeFile = Join-Path $RepoPath ".git\info\exclude"
    if (-not (Test-Path $excludeFile)) {
        return
    }

    $wanted = @(
        "node_modules/",
        "**/node_modules/"
    )
    $existing = Get-Content -Path $excludeFile -ErrorAction SilentlyContinue

    foreach ($pattern in $wanted) {
        if (-not ($existing -contains $pattern)) {
            Add-Content -Path $excludeFile -Value $pattern
        }
    }
}

function Invoke-ClaspPull {
    param(
        [string]$RepoPath,
        [string]$ClaspCommand
    )

    if (-not (Test-Path $ClaspCommand)) {
        Write-Host "[WARN] clasp.cmd not found. Skipping clasp pull."
        return
    }

    $claspDirs = Get-ChildItem -Path $RepoPath -Filter ".clasp.json" -File -Recurse -ErrorAction SilentlyContinue |
        ForEach-Object { $_.DirectoryName } |
        Sort-Object -Unique

    foreach ($dir in $claspDirs) {
        Write-Host "[INFO] clasp pull in: $dir"
        Push-Location $dir
        try {
            $claspOut = cmd /c ('"' + $ClaspCommand + '" pull') 2>&1
            if ($LASTEXITCODE -ne 0) {
                $joined = ($claspOut | Out-String).Trim()
                Write-Host "[WARN] clasp pull failed in ${dir}. ExitCode=$LASTEXITCODE. $joined"
            }
        }
        catch {
            Write-Host "[WARN] clasp pull failed in ${dir}: $($_.Exception.Message)"
        }
        finally {
            Pop-Location
        }
    }
}

try {
    Write-Host "[INFO] Mode: $Mode"
    Write-Host "[INFO] WorkspaceRoot: $WorkspaceRoot"
    Write-Host "[INFO] ProjectsRoot: $ProjectsRoot"
    Write-Host "[INFO] gitCmd: $gitCmd"
    Write-Host "[INFO] claspCmd: $claspCmd"

    $repos = Get-GitRepos -BasePath $ProjectsRoot

    if (-not $repos -or $repos.Count -eq 0) {
        Write-Host "[INFO] No git repos found under $ProjectsRoot"
        exit 0
    }

    $synced = 0
    $changed = 0
    $failed = 0

    foreach ($repo in $repos) {
        Write-Host "`n===================================================="
        Write-Host "[INFO] Repo: $repo"

        try {
            Ensure-LocalExclude -RepoPath $repo
            Invoke-ClaspPull -RepoPath $repo -ClaspCommand $claspCmd

            $remoteText = Get-GitText (Invoke-Git -RepoPath $repo -GitArgs @("remote"))
            $hasOrigin = ($remoteText -split "`r?`n" | Where-Object { $_ -eq "origin" }).Count -gt 0

            $branch = Get-GitText (Invoke-Git -RepoPath $repo -GitArgs @("branch","--show-current"))
            $upstreamProbe = Invoke-Git -RepoPath $repo -GitArgs @("rev-parse","--abbrev-ref","--symbolic-full-name","@{u}") -AllowedExitCodes @(0,128)
            $hasUpstream = $upstreamProbe.ExitCode -eq 0

            if ($hasOrigin -and $hasUpstream) {
                Write-Host "[INFO] git pull --rebase --autostash"
                Invoke-Git -RepoPath $repo -GitArgs @("pull","--rebase","--autostash") | Out-Null
            }
            elseif ($hasOrigin) {
                Write-Host "[INFO] No upstream configured. Running git fetch origin."
                Invoke-Git -RepoPath $repo -GitArgs @("fetch","origin") | Out-Null
            }
            else {
                Write-Host "[WARN] No origin remote found. Backup will stay local for this repo."
            }

            $status = Get-GitText (Invoke-Git -RepoPath $repo -GitArgs @("status","--porcelain"))
            if (-not [string]::IsNullOrWhiteSpace($status)) {
                Write-Host "[INFO] Changes detected. Committing..."
                Invoke-Git -RepoPath $repo -GitArgs @("add","-A") | Out-Null

                $postAddStatus = Get-GitText (Invoke-Git -RepoPath $repo -GitArgs @("status","--porcelain"))
                if (-not [string]::IsNullOrWhiteSpace($postAddStatus)) {
                    $msg = "auto-sync " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + " [" + $env:COMPUTERNAME + "]"
                    Invoke-Git -RepoPath $repo -GitArgs @("commit","-m",$msg) | Out-Null

                    if ($hasOrigin) {
                        if ($hasUpstream) {
                            Invoke-Git -RepoPath $repo -GitArgs @("push") | Out-Null
                        }
                        elseif (-not [string]::IsNullOrWhiteSpace($branch)) {
                            Invoke-Git -RepoPath $repo -GitArgs @("push","-u","origin",$branch) | Out-Null
                        }
                        else {
                            Write-Host "[WARN] Cannot detect current branch. Commit created but not pushed."
                        }
                    }
                    $changed++
                }
                else {
                    Write-Host "[INFO] Only ignored changes found after add. No commit created."
                }
            }
            else {
                Write-Host "[INFO] No changes to commit."
            }

            $synced++
        }
        catch {
            $failed++
            Write-Host "[ERROR] Sync failed in $repo"
            Write-Host "[ERROR] $($_.Exception.Message)"
        }
    }

    Write-Host "`n====================== SUMMARY ======================"
    Write-Host "[INFO] Repos detected: $($repos.Count)"
    Write-Host "[INFO] Repos synced:   $synced"
    Write-Host "[INFO] Repos changed:  $changed"
    Write-Host "[INFO] Repos failed:   $failed"
    Write-Host "[INFO] Log file:       $logFile"
}
finally {
    Stop-Transcript | Out-Null
    if (Test-Path $lockFile) {
        Remove-Item -Path $lockFile -Force
    }
}

