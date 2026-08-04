$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

if (-not (Test-Path '.\.venv\Scripts\python.exe')) {
    throw '.venv not found in project root.'
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$stagingRoot = Join-Path $PSScriptRoot "build_staging\$timestamp"
$stagingWork = Join-Path $stagingRoot 'work'
$stagingDist = Join-Path $stagingRoot 'dist'
$stagedExe = Join-Path $stagingDist 'WB_INN_Extractor.exe'
$targetExe = Join-Path $PSScriptRoot 'dist\WB_INN_Extractor.exe'
$pendingExe = Join-Path $PSScriptRoot "dist\WB_INN_Extractor.$timestamp.new.exe"
$backupDir = Join-Path $PSScriptRoot 'dist\_exe_backups'
$backupExe = Join-Path $backupDir "WB_INN_Extractor.$timestamp.exe"

& .\.venv\Scripts\python.exe -m pip install -r requirements-exe.txt
& .\.venv\Scripts\python.exe -m PyInstaller `
    --noconfirm `
    --clean `
    --workpath $stagingWork `
    --distpath $stagingDist `
    wb_inn_gui.spec

if (-not (Test-Path -LiteralPath $stagedExe -PathType Leaf)) {
    throw "Staged executable was not created: $stagedExe"
}

$pythonDllEntry = & .\.venv\Scripts\pyi-archive_viewer.exe -l $stagedExe |
    Select-String -SimpleMatch 'python314.dll'
if (-not $pythonDllEntry) {
    throw 'Staged executable is invalid: python314.dll is missing from its archive.'
}

$runningTarget = @(
    Get-Process -Name 'WB_INN_Extractor' -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -eq $targetExe }
)
if ($runningTarget.Count -gt 0) {
    throw 'Close WB_INN_Extractor before installing the new build.'
}

New-Item -ItemType Directory -Path (Split-Path -Parent $targetExe) -Force | Out-Null
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
Copy-Item -LiteralPath $stagedExe -Destination $pendingExe

$stagedHash = (Get-FileHash -LiteralPath $stagedExe -Algorithm SHA256).Hash
$pendingHash = (Get-FileHash -LiteralPath $pendingExe -Algorithm SHA256).Hash
if ($stagedHash -ne $pendingHash) {
    throw 'Staged executable copy failed hash verification.'
}

$targetWasMoved = $false
try {
    if (Test-Path -LiteralPath $targetExe -PathType Leaf) {
        Move-Item -LiteralPath $targetExe -Destination $backupExe
        $targetWasMoved = $true
    }
    Move-Item -LiteralPath $pendingExe -Destination $targetExe
}
catch {
    if ($targetWasMoved -and -not (Test-Path -LiteralPath $targetExe)) {
        Move-Item -LiteralPath $backupExe -Destination $targetExe
    }
    throw
}

$installedHash = (Get-FileHash -LiteralPath $targetExe -Algorithm SHA256).Hash
if ($installedHash -ne $stagedHash) {
    throw 'Installed executable failed hash verification.'
}

Write-Host "Build installed: $targetExe"
if ($targetWasMoved) {
    Write-Host "Previous build backed up: $backupExe"
}
