param(
    [ValidateRange(1, 1440)]
    [int]$IntervalMinutes = 10,
    [string]$WorkspaceRoot = "C:\Users\elrub\Desktop\CARPETA CODEX",
    [string]$TaskNameBase = "Codex-AutoSync-GitHub"
)

$ErrorActionPreference = "Stop"

$syncScript = Join-Path $WorkspaceRoot "06_AUTOMATIZACION\sync-all.ps1"
if (-not (Test-Path $syncScript)) {
    throw "sync-all.ps1 not found at: $syncScript"
}

$psExe = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
$taskCmd = '"' + $psExe + '" -NoProfile -ExecutionPolicy Bypass -File "' + $syncScript + '" -Mode backup -WorkspaceRoot "' + $WorkspaceRoot + '"'

$taskPeriodic = "\$TaskNameBase-Periodic"
$taskLogon = "\$TaskNameBase-AtLogon"

function Remove-TaskIfExists {
    param([string]$TaskName)

    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        & schtasks.exe /Query /TN $TaskName > $null 2>&1
        if ($LASTEXITCODE -eq 0) {
            & schtasks.exe /Delete /TN $TaskName /F > $null 2>&1
        }
    }
    finally {
        $ErrorActionPreference = $prev
    }
}

function Create-TaskOrThrow {
    param([string[]]$Args, [string]$TaskName)

    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $out = & schtasks.exe @Args 2>&1
        $code = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $prev
    }

    if ($code -ne 0) {
        $txt = ($out | Out-String).Trim()
        throw "Failed creating task $TaskName. ExitCode=$code. $txt"
    }
}

Remove-TaskIfExists -TaskName $taskPeriodic
Remove-TaskIfExists -TaskName $taskLogon

Create-TaskOrThrow -TaskName $taskPeriodic -Args @(
    "/Create", "/TN", $taskPeriodic, "/SC", "MINUTE", "/MO", "$IntervalMinutes", "/TR", $taskCmd, "/F"
)

Create-TaskOrThrow -TaskName $taskLogon -Args @(
    "/Create", "/TN", $taskLogon, "/SC", "ONLOGON", "/TR", $taskCmd, "/F"
)

Write-Host "Task created: $taskPeriodic (every $IntervalMinutes minutes)"
Write-Host "Task created: $taskLogon (at logon)"
Write-Host "Command: $taskCmd"
