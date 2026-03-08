param(
    [string]$TaskNameBase = "Codex-AutoSync-GitHub"
)

$ErrorActionPreference = "Stop"
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
            Write-Host "Removed task: $TaskName"
        }
        else {
            Write-Host "Task not found (skip): $TaskName"
        }
    }
    finally {
        $ErrorActionPreference = $prev
    }
}

Remove-TaskIfExists -TaskName $taskPeriodic
Remove-TaskIfExists -TaskName $taskLogon
