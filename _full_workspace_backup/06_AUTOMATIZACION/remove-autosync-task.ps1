param(
    [string]$TaskName = "Codex-AutoSync-GitHub"
)

$ErrorActionPreference = "Continue"

$targets = @(
    $TaskName,
    "$TaskName-Periodic",
    "$TaskName-AtLogon"
)

foreach ($t in $targets) {
    schtasks /Query /TN $t > $null 2>&1
    if ($LASTEXITCODE -eq 0) {
        schtasks /Delete /TN $t /F > $null 2>&1
        Write-Host "Removed task: \\$t"
    }
    else {
        Write-Host "Task not found (skip): \\$t"
    }
}
