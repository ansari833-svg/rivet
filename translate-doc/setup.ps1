# setup.ps1 — install translate-doc on Windows.
#
# Run with:
#   powershell -ExecutionPolicy Bypass -File setup.ps1
#
# Creates a virtual environment, installs the tool into it, drops a launcher
# batch file, and adds that launcher's folder to your user PATH.

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

Write-Host "Creating virtual environment (.venv)..."
python -m venv "$root\.venv"

Write-Host "Installing dependencies..."
& "$root\.venv\Scripts\pip" install -e "$root"

$venvPython = Join-Path $root ".venv\Scripts\python.exe"

# --- Launcher batch file -----------------------------------------------------
# Invokes the venv's Python directly so the tool works without activating the
# venv first.
$launcherDir = Join-Path $env:LOCALAPPDATA "Programs\translate-doc"
if (-not (Test-Path $launcherDir)) {
    New-Item -ItemType Directory -Path $launcherDir -Force | Out-Null
}
$launcherPath = Join-Path $launcherDir "translate-doc.bat"
$batContent = "@echo off`r`n`"$venvPython`" -m translate_doc.main %*`r`n"
Set-Content -Path $launcherPath -Value $batContent -Encoding ASCII
Write-Host "Launcher created at $launcherPath"

# --- Add launcher folder to the user PATH ------------------------------------
# NOTE: never use `setx PATH` — it truncates at 1024 characters and can destroy
# the existing PATH. Read the current value, check, then append.
$userPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
if ($null -eq $userPath) { $userPath = '' }

$existing = $userPath -split ';' | Where-Object { $_ -ne '' }
if ($existing -contains $launcherDir) {
    Write-Host "$launcherDir is already on your user PATH."
} else {
    $trimmed = $userPath.TrimEnd(';')
    if ($trimmed -eq '') {
        $newPath = $launcherDir
    } else {
        $newPath = "$trimmed;$launcherDir"
    }
    [Environment]::SetEnvironmentVariable('PATH', $newPath, 'User')
    Write-Host "Added $launcherDir to your user PATH."
}

Write-Host ""
Write-Host "Setup complete."
Write-Host "IMPORTANT: open a NEW terminal window for the PATH change to take effect."
Write-Host "Then run:  translate-doc"
