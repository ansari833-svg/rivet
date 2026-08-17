# setup.ps1 - Windows setup for arabic-ocr
#
# Run with:
#   powershell -ExecutionPolicy Bypass -File setup.ps1
#
# This installs a CPU-only PyTorch first (so surya-ocr does not pull several GB
# of CUDA libraries on a machine without an NVIDIA GPU), then installs the tool,
# creates a launcher, and puts it on the user PATH.

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $scriptDir

$venvDir = Join-Path $scriptDir ".venv"
$venvPython = Join-Path $venvDir "Scripts\python.exe"
$venvPip = Join-Path $venvDir "Scripts\pip.exe"

# --- Virtual environment ---------------------------------------------------- #
if (-not (Test-Path $venvPython)) {
    Write-Host "Creating virtual environment at .venv ..." -ForegroundColor Cyan
    python -m venv .venv
} else {
    Write-Host "Virtual environment already exists - reusing .venv" -ForegroundColor Cyan
}

Write-Host "Upgrading pip ..." -ForegroundColor Cyan
& $venvPython -m pip install --upgrade pip

# --- CPU PyTorch first ------------------------------------------------------ #
# Install the CPU-only wheel explicitly BEFORE surya-ocr, so surya's dependency
# resolution does not download the CUDA build.
Write-Host "Installing CPU-only PyTorch ..." -ForegroundColor Cyan
& $venvPip install torch --index-url https://download.pytorch.org/whl/cpu

# --- The tool --------------------------------------------------------------- #
Write-Host "Installing arabic-ocr and dependencies ..." -ForegroundColor Cyan
& $venvPip install -e .

# --- NVIDIA GPU note (informational only) ----------------------------------- #
$hasGpu = $false
try {
    nvidia-smi | Out-Null
    if ($LASTEXITCODE -eq 0) { $hasGpu = $true }
} catch {
    $hasGpu = $false
}
if ($hasGpu) {
    Write-Host ""
    Write-Host "An NVIDIA GPU was detected." -ForegroundColor Yellow
    Write-Host "The CUDA build of PyTorch would give a 5-10x speedup." -ForegroundColor Yellow
    Write-Host "It is NOT installed automatically. To install it, run:" -ForegroundColor Yellow
    Write-Host "  .venv\Scripts\pip install torch --index-url https://download.pytorch.org/whl/cu121" -ForegroundColor Yellow
    Write-Host ""
}

# --- Launcher batch file ---------------------------------------------------- #
$launcherDir = Join-Path $env:LOCALAPPDATA "Programs\arabic-ocr"
if (-not (Test-Path $launcherDir)) {
    New-Item -ItemType Directory -Path $launcherDir -Force | Out-Null
}
$launcherPath = Join-Path $launcherDir "arabic-ocr.bat"
# Invoke the venv's Python directly so the tool works from any terminal.
$launcherBody = "@echo off`r`n`"$venvPython`" -m arabic_ocr.main %*`r`n"
Set-Content -Path $launcherPath -Value $launcherBody -Encoding ASCII
Write-Host "Launcher written to $launcherPath" -ForegroundColor Cyan

# --- User PATH -------------------------------------------------------------- #
# Read the current user PATH first, check the folder isn't already present, then
# append. Do NOT use `setx PATH` - it truncates at 1024 characters and can
# destroy the existing PATH.
$currentPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
if ($null -eq $currentPath) { $currentPath = "" }

$pathEntries = $currentPath.Split(';') | Where-Object { $_ -ne "" }
$alreadyPresent = $pathEntries | Where-Object { $_.TrimEnd('\') -ieq $launcherDir.TrimEnd('\') }

if (-not $alreadyPresent) {
    if ($currentPath -eq "") {
        $newPath = $launcherDir
    } else {
        $newPath = $currentPath.TrimEnd(';') + ";" + $launcherDir
    }
    [Environment]::SetEnvironmentVariable('PATH', $newPath, 'User')
    Write-Host "Added $launcherDir to your user PATH." -ForegroundColor Green
} else {
    Write-Host "$launcherDir is already on your user PATH." -ForegroundColor Green
}

Write-Host ""
Write-Host "Setup complete." -ForegroundColor Green
Write-Host "IMPORTANT: open a NEW terminal window for the PATH change to take effect." -ForegroundColor Yellow
Write-Host "Then run:  arabic-ocr" -ForegroundColor Yellow
