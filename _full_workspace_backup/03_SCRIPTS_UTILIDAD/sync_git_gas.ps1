param(
  [ValidateSet("status", "down", "up", "backup", "audit")]
  [string]$Action = "status",
  [string]$RepoPath = "C:\Users\elrub\Desktop\CARPETA CODEX\01_PROYECTOS\festivales-github",
  [string]$Branch = "main",
  [string]$CommitMessage = "",
  [string]$DriveBackupPath = "",
  [string]$AuditOutputPath = "",
  [switch]$SkipClasp,
  [switch]$AllowDirty,
  [switch]$SkipNetworkChecks
)

$ErrorActionPreference = "Stop"
$script:GitExe = $null
$script:NodeExe = $null
$script:WorkspaceRoot = Split-Path (Split-Path $RepoPath -Parent) -Parent
$script:AuditFindings = @()

function Write-Step {
  param([string]$Message)
  Write-Host "[SYNC] $Message"
}

function Add-Finding {
  param(
    [string]$Severity,
    [string]$Check,
    [string]$Status,
    [string]$Detail
  )
  $script:AuditFindings += [pscustomobject]@{
    Severity = $Severity
    Check = $Check
    Status = $Status
    Detail = $Detail
  }
}

function Resolve-Exe {
  param(
    [string]$Name,
    [string[]]$Candidates
  )

  $cmd = Get-Command $Name -ErrorAction SilentlyContinue
  if ($cmd -and $cmd.Source) {
    return $cmd.Source
  }

  foreach ($candidate in $Candidates) {
    if (Test-Path -LiteralPath $candidate) {
      return $candidate
    }
  }

  throw "No se encontro '$Name' ni en PATH ni en rutas candidatas."
}

function Ensure-Repo {
  if (-not (Test-Path -LiteralPath $RepoPath)) {
    throw "RepoPath no existe: $RepoPath"
  }
  if (-not (Test-Path -LiteralPath (Join-Path $RepoPath ".git"))) {
    throw "No se detecta .git en: $RepoPath"
  }
}

function Resolve-Tools {
  $script:GitExe = Resolve-Exe -Name "git" -Candidates @(
    "C:\Program Files\Git\cmd\git.exe",
    "C:\Program Files (x86)\Git\cmd\git.exe"
  )

  $script:NodeExe = Resolve-Exe -Name "node" -Candidates @(
    "C:\Program Files\nodejs\node.exe",
    "C:\Program Files (x86)\nodejs\node.exe"
  )

  $gitDir = Split-Path -Parent $script:GitExe
  $nodeDir = Split-Path -Parent $script:NodeExe
  $env:Path = "$gitDir;$nodeDir;$env:Path"
}

function Get-ClaspLocalPath {
  return (Join-Path $RepoPath "node_modules\@google\clasp\build\src\index.js")
}

function Invoke-Git {
  param(
    [string[]]$GitArgs,
    [switch]$AllowFailure
  )

  & $script:GitExe -c "safe.directory=$RepoPath" @GitArgs
  $code = $LASTEXITCODE
  if (-not $AllowFailure -and $code -ne 0) {
    throw "Error de git: $($GitArgs -join ' ')"
  }
  return $code
}

function Invoke-Clasp {
  param(
    [string[]]$ClaspArgs,
    [switch]$AllowFailure
  )

  $claspLocal = Get-ClaspLocalPath
  if (-not (Test-Path -LiteralPath $claspLocal)) {
    if ($AllowFailure) {
      return 127
    }
    throw "No se encontro clasp local en node_modules. Ejecuta npm install en el repo."
  }

  & $script:NodeExe $claspLocal @ClaspArgs
  $code = $LASTEXITCODE
  if (-not $AllowFailure -and $code -ne 0) {
    throw "Error ejecutando clasp: $($ClaspArgs -join ' ')"
  }
  return $code
}

function Git-IsDirty {
  $out = & $script:GitExe -c "safe.directory=$RepoPath" status --porcelain
  return -not [string]::IsNullOrWhiteSpace(($out | Out-String))
}

function Git-HasStagedChanges {
  & $script:GitExe -c "safe.directory=$RepoPath" diff --cached --quiet
  return ($LASTEXITCODE -ne 0)
}

function Run-Status {
  Write-Step "Estado de Git"
  Invoke-Git -GitArgs @("status", "--short", "--branch") | Out-Null
  if (-not $SkipClasp) {
    Write-Step "Estado de Apps Script (clasp)"
    Invoke-Clasp -ClaspArgs @("status") | Out-Null
  }
}

function Run-Down {
  if ((Git-IsDirty) -and (-not $AllowDirty)) {
    throw "Repo con cambios locales. Usa -AllowDirty o limpia el arbol antes de hacer down."
  }

  Write-Step "Descargando cambios de GitHub ($Branch)"
  Invoke-Git -GitArgs @("pull", "--ff-only", "origin", $Branch) | Out-Null

  if (-not $SkipClasp) {
    Write-Step "Descargando cambios de Google Apps Script"
    Invoke-Clasp -ClaspArgs @("pull") | Out-Null
  }

  Write-Step "Sincronizacion DOWN completada"
}

function Run-Up {
  if (-not $SkipClasp) {
    Write-Step "Subiendo cambios a Google Apps Script"
    Invoke-Clasp -ClaspArgs @("push") | Out-Null
  }

  Write-Step "Preparando commit en Git"
  Invoke-Git -GitArgs @("add", "-A") | Out-Null

  if (-not (Git-HasStagedChanges)) {
    Write-Step "No hay cambios para commit. Se omite push a GitHub."
    return
  }

  $finalMessage = $CommitMessage
  if ([string]::IsNullOrWhiteSpace($finalMessage)) {
    $finalMessage = "sync: " + (Get-Date -Format "yyyy-MM-dd HH:mm")
  }

  Invoke-Git -GitArgs @("commit", "-m", $finalMessage) | Out-Null
  Write-Step "Subiendo commit a GitHub ($Branch)"
  Invoke-Git -GitArgs @("push", "origin", $Branch) | Out-Null
  Write-Step "Sincronizacion UP completada"
}

function Run-Backup {
  if ([string]::IsNullOrWhiteSpace($DriveBackupPath)) {
    throw "Para backup debes indicar -DriveBackupPath."
  }
  if (-not (Test-Path -LiteralPath $DriveBackupPath)) {
    throw "No existe DriveBackupPath: $DriveBackupPath"
  }

  $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
  $repoName = Split-Path -Path $RepoPath -Leaf
  $dest = Join-Path $DriveBackupPath ($repoName + "_backup_" + $stamp)

  Write-Step "Creando backup one-way en $dest"
  New-Item -ItemType Directory -Path $dest -Force | Out-Null
  & robocopy $RepoPath $dest /E /XD .git node_modules /XF ".clasp.json"
  $code = $LASTEXITCODE
  if ($code -gt 7) {
    throw "Robocopy fallo con codigo $code"
  }

  Write-Step "Backup completado"
}

function Run-Audit {
  Write-Step "Iniciando auditoria completa"
  $script:AuditFindings = @()

  Add-Finding -Severity "INFO" -Check "RepoPath" -Status "OK" -Detail $RepoPath

  foreach ($required in @("package.json", ".clasp.json", ".gitignore", ".claspignore")) {
    $full = Join-Path $RepoPath $required
    if (Test-Path -LiteralPath $full) {
      Add-Finding -Severity "INFO" -Check "Archivo requerido" -Status "OK" -Detail $required
    }
    else {
      Add-Finding -Severity "ERROR" -Check "Archivo requerido" -Status "MISSING" -Detail $required
    }
  }

  $gitStatusRaw = & $script:GitExe -c "safe.directory=$RepoPath" status --short --branch 2>&1
  $gitStatus = ($gitStatusRaw | Out-String).Trim()
  Add-Finding -Severity "INFO" -Check "Git status" -Status "OK" -Detail ($gitStatus -replace "`r?`n", " <br> ")

  if (Git-IsDirty) {
    Add-Finding -Severity "WARN" -Check "Working tree" -Status "DIRTY" -Detail "Hay cambios locales o no trackeados."
  }
  else {
    Add-Finding -Severity "INFO" -Check "Working tree" -Status "CLEAN" -Detail "Sin cambios locales."
  }

  $remoteRaw = & $script:GitExe -c "safe.directory=$RepoPath" remote -v 2>&1
  $remoteText = ($remoteRaw | Out-String).Trim()
  Add-Finding -Severity "INFO" -Check "Git remote" -Status "OK" -Detail ($remoteText -replace "`r?`n", " <br> ")

  if (-not $SkipNetworkChecks) {
    $null = Invoke-Git -GitArgs @("ls-remote", "--heads", "origin") -AllowFailure
    if ($LASTEXITCODE -ne 0) {
      Add-Finding -Severity "WARN" -Check "Git network" -Status "FAIL" -Detail "No se pudo validar conectividad a origin."
    }
    else {
      Add-Finding -Severity "INFO" -Check "Git network" -Status "OK" -Detail "Conectividad con origin confirmada."
    }
  }

  $claspLocal = Get-ClaspLocalPath
  if (-not (Test-Path -LiteralPath $claspLocal)) {
    Add-Finding -Severity "ERROR" -Check "clasp local" -Status "MISSING" -Detail "Falta node_modules/@google/clasp."
  }
  elseif ($SkipClasp) {
    Add-Finding -Severity "INFO" -Check "clasp checks" -Status "SKIPPED" -Detail "Se omitio por -SkipClasp."
  }
  else {
    $authRaw = & $script:NodeExe $claspLocal show-authorized-user 2>&1
    $authText = ($authRaw | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) {
      Add-Finding -Severity "ERROR" -Check "clasp auth" -Status "FAIL" -Detail ($authText -replace "`r?`n", " <br> ")
    }
    elseif ($authText -match "unknown user") {
      Add-Finding -Severity "WARN" -Check "clasp auth" -Status "UNKNOWN" -Detail "Autenticado sin usuario claro. Recomendado relogin."
    }
    else {
      Add-Finding -Severity "INFO" -Check "clasp auth" -Status "OK" -Detail ($authText -replace "`r?`n", " <br> ")
    }

    $statusRaw = & $script:NodeExe $claspLocal status 2>&1
    $statusText = ($statusRaw | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) {
      Add-Finding -Severity "ERROR" -Check "clasp status" -Status "FAIL" -Detail ($statusText -replace "`r?`n", " <br> ")
    }
    else {
      Add-Finding -Severity "INFO" -Check "clasp status" -Status "OK" -Detail "Comando ejecutado sin error."
    }
  }

  $tmpFile = Join-Path $RepoPath "Codigo_tmp.js"
  if (Test-Path -LiteralPath $tmpFile) {
    Add-Finding -Severity "WARN" -Check "Archivo temporal" -Status "PRESENT" -Detail "Existe Codigo_tmp.js; puede generar confusion de fuente de verdad."
  }

  $mainClasp = (Join-Path $RepoPath ".clasp.json")
  $nestedClasp = Get-ChildItem -Path $RepoPath -Recurse -Force -Filter ".clasp.json" |
    Where-Object { $_.FullName -ne $mainClasp }

  if ($nestedClasp.Count -gt 0) {
    $paths = ($nestedClasp | ForEach-Object { $_.FullName }) -join " | "
    Add-Finding -Severity "WARN" -Check "Multiples .clasp.json" -Status "PRESENT" -Detail $paths
  }

  $gitignorePath = Join-Path $RepoPath ".gitignore"
  if (Test-Path -LiteralPath $gitignorePath) {
    $gitignore = Get-Content -Raw $gitignorePath
    if ($gitignore -notmatch "node_modules/") {
      Add-Finding -Severity "WARN" -Check ".gitignore" -Status "MISSING_RULE" -Detail "Falta regla node_modules/."
    }
    if ($gitignore -notmatch "\.clasprc\.json") {
      Add-Finding -Severity "WARN" -Check ".gitignore" -Status "MISSING_RULE" -Detail "Falta regla .clasprc.json."
    }
  }

  if ([string]::IsNullOrWhiteSpace($AuditOutputPath)) {
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $AuditOutputPath = Join-Path $script:WorkspaceRoot ("05_DOCUMENTACION\\AUDITORIA_SYNC_" + $stamp + ".md")
  }

  $reportDir = Split-Path -Parent $AuditOutputPath
  if (-not (Test-Path -LiteralPath $reportDir)) {
    New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
  }

  $okCount = ($script:AuditFindings | Where-Object { $_.Severity -eq "INFO" }).Count
  $warnCount = ($script:AuditFindings | Where-Object { $_.Severity -eq "WARN" }).Count
  $errCount = ($script:AuditFindings | Where-Object { $_.Severity -eq "ERROR" }).Count

  $lines = @()
  $lines += "# Auditoria De Sincronizacion"
  $lines += ""
  $lines += "Fecha: " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
  $lines += "Repo: $RepoPath"
  $lines += ""
  $lines += "Resumen:"
  $lines += "- INFO: $okCount"
  $lines += "- WARN: $warnCount"
  $lines += "- ERROR: $errCount"
  $lines += ""
  $lines += "| Severity | Check | Status | Detail |"
  $lines += "|---|---|---|---|"

  foreach ($f in $script:AuditFindings) {
    $detail = ($f.Detail -replace "\|", "\\|") -replace "`r?`n", " <br> "
    $lines += "| $($f.Severity) | $($f.Check) | $($f.Status) | $detail |"
  }

  Set-Content -Encoding UTF8 -Path $AuditOutputPath -Value $lines
  Write-Step "Auditoria guardada en: $AuditOutputPath"

  if ($errCount -gt 0) {
    throw "Auditoria completada con errores criticos. Revisa el reporte."
  }
}

try {
  Ensure-Repo
  Resolve-Tools

  Push-Location $RepoPath
  switch ($Action) {
    "status" { Run-Status }
    "down" { Run-Down }
    "up" { Run-Up }
    "backup" { Run-Backup }
    "audit" { Run-Audit }
  }
}
finally {
  Pop-Location -ErrorAction SilentlyContinue
}
