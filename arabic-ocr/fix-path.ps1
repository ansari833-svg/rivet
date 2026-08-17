# fix-path.ps1 - diagnose and repair the arabic-ocr launcher + PATH entry
#
# Run from the arabic-ocr project folder (the one containing .venv):
#   powershell -ExecutionPolicy Bypass -File fix-path.ps1
#
# It locates the venv python, checks whether the launcher .bat exists and
# whether its folder is on the user PATH, fixes whichever step failed, and
# refreshes the CURRENT shell so you can verify without opening a new window.

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

$venvPython  = Join-Path $scriptDir ".venv\Scripts\python.exe"
$launcherDir = Join-Path $env:LOCALAPPDATA "Programs\arabic-ocr"
$launcherPath = Join-Path $launcherDir "arabic-ocr.bat"

Write-Host "=== Diagnostics ===" -ForegroundColor Cyan
Write-Host "Project folder : $scriptDir"
Write-Host "Venv python    : $venvPython"
Write-Host "Launcher folder: $launcherDir"
Write-Host "Launcher .bat  : $launcherPath"
Write-Host ""

# --- 1. Locate the venv python --------------------------------------------- #
if (-not (Test-Path $venvPython)) {
    Write-Host "[fix] venv python not at the expected path - searching..." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Contents of $scriptDir :" -ForegroundColor Cyan
    Get-ChildItem -Force $scriptDir | ForEach-Object {
        $tag = if ($_.PSIsContainer) { '<dir> ' } else { '      ' }
        Write-Host "  $tag$($_.Name)"
    }

    $all = Get-ChildItem -Path $scriptDir -Recurse -Depth 4 -Force -Filter python.exe -ErrorAction SilentlyContinue
    $found = @($all | Where-Object { $_.FullName -match '[\\/]Scripts[\\/]python\.exe$' })

    if ($found.Count -ge 1) {
        $venvPython = $found[0].FullName
        Write-Host "[ok] found a venv python at: $venvPython" -ForegroundColor Green
        if ($found.Count -gt 1) {
            Write-Host "     (multiple found; using the first - all listed below)" -ForegroundColor Yellow
            $found | ForEach-Object { Write-Host "       $($_.FullName)" }
        }
    } else {
        Write-Host "PROBLEM: no .venv\Scripts\python.exe found under this folder." -ForegroundColor Red
        Write-Host ""
        Write-Host "The virtual environment was never created here. Make sure this folder" -ForegroundColor Red
        Write-Host "contains pyproject.toml and the arabic_ocr\ package, then from THIS" -ForegroundColor Red
        Write-Host "folder run setup first:" -ForegroundColor Red
        Write-Host "  powershell -ExecutionPolicy Bypass -File setup.ps1" -ForegroundColor Red
        Write-Host "and re-run this script afterward." -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "[ok] venv python exists" -ForegroundColor Green
}

# --- 2. Launcher .bat ------------------------------------------------------- #
$batOk = $false
if (Test-Path $launcherPath) {
    $existing = Get-Content $launcherPath -Raw
    if ($existing -like "*$venvPython*") {
        Write-Host "[ok] launcher .bat exists and points at this venv" -ForegroundColor Green
        $batOk = $true
    } else {
        Write-Host "[fix] launcher .bat exists but points elsewhere - rewriting" -ForegroundColor Yellow
    }
} else {
    Write-Host "[fix] launcher .bat missing - creating it" -ForegroundColor Yellow
}

if (-not $batOk) {
    if (-not (Test-Path $launcherDir)) {
        New-Item -ItemType Directory -Path $launcherDir -Force | Out-Null
    }
    $launcherBody = "@echo off`r`n`"$venvPython`" -m arabic_ocr.main %*`r`n"
    Set-Content -Path $launcherPath -Value $launcherBody -Encoding ASCII
    Write-Host "[done] wrote $launcherPath" -ForegroundColor Green
}

# --- 3. User PATH ----------------------------------------------------------- #
# Read the current USER PATH first, then append only if the folder is absent.
# Never use setx - it truncates at 1024 chars and can destroy the existing PATH.
$currentPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
if ($null -eq $currentPath) { $currentPath = "" }

$pathEntries = $currentPath.Split(';') | Where-Object { $_ -ne "" }
$alreadyPresent = $pathEntries | Where-Object { $_.TrimEnd('\') -ieq $launcherDir.TrimEnd('\') }

if ($alreadyPresent) {
    Write-Host "[ok] launcher folder is already on the user PATH" -ForegroundColor Green
} else {
    if ($currentPath -eq "") {
        $newPath = $launcherDir
    } else {
        $newPath = $currentPath.TrimEnd(';') + ";" + $launcherDir
    }
    [Environment]::SetEnvironmentVariable('PATH', $newPath, 'User')
    Write-Host "[done] appended launcher folder to the user PATH" -ForegroundColor Green
}

# --- 4. Refresh THIS session ------------------------------------------------ #
# A "new window" spawned from an existing Windows Terminal inherits the old
# environment block, which is the usual reason the command still isn't found.
# Rebuild this process's PATH from the machine + user values so you can test now.
$machinePath = [Environment]::GetEnvironmentVariable('PATH', 'Machine')
$userPath    = [Environment]::GetEnvironmentVariable('PATH', 'User')
$env:PATH = (@($machinePath, $userPath) | Where-Object { $_ } ) -join ';'

# --- 5. Verify -------------------------------------------------------------- #
Write-Host ""
Write-Host "=== Verify ===" -ForegroundColor Cyan
$resolved = Get-Command arabic-ocr -ErrorAction SilentlyContinue
if ($resolved) {
    Write-Host "[ok] 'arabic-ocr' resolves to: $($resolved.Source)" -ForegroundColor Green
    Write-Host ""
    Write-Host "Fixed. It works in THIS window now." -ForegroundColor Green
    Write-Host "If a brand-new window still can't find it, fully CLOSE all terminal" -ForegroundColor Yellow
    Write-Host "windows (Windows Terminal caches its environment) and open a fresh one." -ForegroundColor Yellow
} else {
    Write-Host "[warn] 'arabic-ocr' still does not resolve in this refreshed session." -ForegroundColor Red
    Write-Host "The launcher is at: $launcherPath" -ForegroundColor Red
    Write-Host "Try running it directly to confirm it works: `"$launcherPath`" --help" -ForegroundColor Red
}
