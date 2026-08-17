# setup.ps1 - Windows setup for the pdf-screenshots CLI tool.
#
# Run with:
#   powershell -ExecutionPolicy Bypass -File setup.ps1
#
# The default execution policy will otherwise block this script.

$ErrorActionPreference = "Stop"

# Resolve paths relative to this script so it works from any working directory.
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $ScriptDir

Write-Host "Creating virtual environment at .venv ..."
python -m venv .venv

$VenvPython = Join-Path $ScriptDir ".venv\Scripts\python.exe"
$VenvPip = Join-Path $ScriptDir ".venv\Scripts\pip.exe"

Write-Host "Installing dependencies ..."
& $VenvPip install -e .

# ---------------------------------------------------------------------------
# Create a launcher batch file that invokes the venv's Python directly, so the
# tool works from any terminal without activating the venv first.
# ---------------------------------------------------------------------------
$LauncherDir = Join-Path $env:LOCALAPPDATA "Programs\pdf-screenshots"
if (-not (Test-Path $LauncherDir)) {
    New-Item -ItemType Directory -Path $LauncherDir -Force | Out-Null
}

$LauncherPath = Join-Path $LauncherDir "pdf-screenshots.bat"
$BatchContent = "@echo off`r`n`"$VenvPython`" -m pdf_screenshots.main %*`r`n"
Set-Content -Path $LauncherPath -Value $BatchContent -Encoding ASCII
Write-Host "Created launcher: $LauncherPath"

# ---------------------------------------------------------------------------
# Add the launcher folder to the user PATH.
#
# Do NOT use `setx PATH` - it truncates at 1024 characters and can destroy the
# existing PATH. Read the current user PATH, check the folder isn't already
# present, then append.
# ---------------------------------------------------------------------------
$currentPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
if ($null -eq $currentPath) {
    $currentPath = ""
}

$pathEntries = $currentPath -split ';' | Where-Object { $_ -ne "" }
if ($pathEntries -notcontains $LauncherDir) {
    if ($currentPath -ne "" -and -not $currentPath.EndsWith(';')) {
        $newPath = "$currentPath;$LauncherDir"
    } else {
        $newPath = "$currentPath$LauncherDir"
    }
    [Environment]::SetEnvironmentVariable('PATH', $newPath, 'User')
    Write-Host "Added $LauncherDir to your user PATH."
} else {
    Write-Host "$LauncherDir is already on your user PATH."
}

Write-Host ""
Write-Host "Setup complete."
Write-Host "IMPORTANT: Open a NEW terminal window for the PATH change to take effect."
Write-Host "Then run:  pdf-screenshots"
