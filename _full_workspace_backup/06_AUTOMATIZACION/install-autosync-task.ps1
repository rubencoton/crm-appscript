param(
    [ValidateRange(1, 1440)]
    [int]$IntervalMinutes = 10,
    [string]$WorkspaceRoot = "C:\Users\elrub\Desktop\CARPETA CODEX",
    [string]$TaskName = "Codex-AutoSync-GitHub"
)

$ErrorActionPreference = "Stop"

$syncScript = Join-Path $WorkspaceRoot "06_AUTOMATIZACION\sync-all.ps1"
if (-not (Test-Path $syncScript)) {
    throw "sync-all.ps1 not found at: $syncScript"
}

$xmlPath = Join-Path $WorkspaceRoot "06_AUTOMATIZACION\autosync-task.xml"
$user = "$env:USERDOMAIN\$env:USERNAME"
$start = (Get-Date).AddMinutes(1).ToString('s')
$interval = "PT${IntervalMinutes}M"

$xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>$((Get-Date).ToString('s'))</Date>
    <Author>$user</Author>
    <Description>Auto backup GitHub from CARPETA CODEX every $IntervalMinutes minutes and at logon.</Description>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>$start</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay>
        <DaysInterval>1</DaysInterval>
      </ScheduleByDay>
      <Repetition>
        <Interval>$interval</Interval>
        <Duration>P1D</Duration>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
    </CalendarTrigger>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>$user</UserId>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$user</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT2H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe</Command>
      <Arguments>-NoProfile -ExecutionPolicy Bypass -File "$syncScript" -Mode backup -WorkspaceRoot "$WorkspaceRoot"</Arguments>
    </Exec>
  </Actions>
</Task>
"@

Set-Content -Path $xmlPath -Value $xml -Encoding Unicode

function Remove-TaskIfExists {
    param([string]$Name)

    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        schtasks /Delete /TN $Name /F 2>$null | Out-Null
    }
    finally {
        $ErrorActionPreference = $prev
    }
}

# Remove current and legacy task variants if present
Remove-TaskIfExists -Name $TaskName
Remove-TaskIfExists -Name "$TaskName-Periodic"
Remove-TaskIfExists -Name "$TaskName-AtLogon"

schtasks /Create /TN $TaskName /XML $xmlPath /F | Out-Null

Write-Host "Task created: \\$TaskName"
Write-Host "Interval: every $IntervalMinutes minutes"
Write-Host "Triggers: periodic + at logon"
Write-Host "Action: powershell sync-all.ps1 -Mode backup"

