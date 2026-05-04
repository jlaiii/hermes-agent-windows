#Requires -Version 5.1
#Requires -RunAsAdministrator

# ==========================================
# Hermes Agent Windows Uninstaller
# Removes: Install directory, shortcuts,
#          Start Menu folder, WSL distro (optional)
# ==========================================

param(
    [switch]$KeepWSL
)

$ErrorActionPreference = "Stop"

# --- Configuration ---
$InstallDir   = "$env:LOCALAPPDATA\HermesAgent"
$LaunchersDir = "$InstallDir\Launchers"
$DistroName   = "Ubuntu-22.04"

# --- Banner ---
Write-Host ""
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "  Hermes Agent - Uninstaller"                         -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host ""

# Confirm
$confirm = Read-Host "This will remove Hermes Agent from your system. Continue? [Y/n]"
if ($confirm -and ($confirm[0] -ne 'Y' -and $confirm[0] -ne 'y')) {
    Write-Host "Uninstall cancelled." -ForegroundColor Yellow
    exit 0
}

# 1. Remove Desktop shortcuts
Write-Host "Removing Desktop shortcuts..." -ForegroundColor Cyan
$Desktop = [Environment]::GetFolderPath("Desktop")
$shortcutNames = @(
    "Hermes CLI.lnk",
    "Hermes Gateway.lnk",
    "Ollama Server.lnk",
    "Hermes Setup.lnk"
)
foreach ($name in $shortcutNames) {
    $path = Join-Path $Desktop $name
    if (Test-Path $path) {
        Remove-Item $path -Force
        Write-Host "  Removed: $path" -ForegroundColor Green
    } else {
        Write-Host "  Not found: $path" -ForegroundColor DarkGray
    }
}

# 2. Remove Start Menu folder
Write-Host "Removing Start Menu folder..." -ForegroundColor Cyan
$StartMenu = [Environment]::GetFolderPath("StartMenu")
$AppFolder = "$StartMenu\Programs\Hermes Agent"
if (Test-Path $AppFolder) {
    Remove-Item $AppFolder -Recurse -Force
    Write-Host "  Removed: $AppFolder" -ForegroundColor Green
} else {
    Write-Host "  Not found: $AppFolder" -ForegroundColor DarkGray
}

# 3. Cleanup RunOnce registry
Write-Host "Cleaning up RunOnce registry entries..." -ForegroundColor Cyan
$regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
try {
    Remove-ItemProperty -Path $regPath -Name "HermesAgentSetup" -ErrorAction SilentlyContinue
    Write-Host "  Registry key removed." -ForegroundColor Green
} catch {
    Write-Host "  No registry key found (or could not remove)." -ForegroundColor DarkGray
}

# 4. Remove install directory
Write-Host "Removing install directory..." -ForegroundColor Cyan
if (Test-Path $InstallDir) {
    Remove-Item $InstallDir -Recurse -Force
    Write-Host "  Removed: $InstallDir" -ForegroundColor Green
} else {
    Write-Host "  Not found: $InstallDir" -ForegroundColor DarkGray
}

# 5. Optionally remove WSL distro
if (-not $KeepWSL) {
    Write-Host "Checking for WSL distro '$DistroName'..." -ForegroundColor Cyan
    try {
        $wslList = wsl -l -v 2>$null
        if ($wslList -match [regex]::Escape($DistroName)) {
            $confirmWSL = Read-Host "  Found '$DistroName'. Unregister and remove it? [Y/n]"
            if (-not $confirmWSL -or ($confirmWSL[0] -eq 'Y' -or $confirmWSL[0] -eq 'y')) {
                Write-Host "  Unregistering $DistroName..." -ForegroundColor Yellow
                wsl --unregister $DistroName 2>&1 | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
                Write-Host "  WSL distro removed." -ForegroundColor Green
            } else {
                Write-Host "  Skipped WSL distro removal." -ForegroundColor Yellow
            }
        } else {
            Write-Host "  WSL distro '$DistroName' not found." -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "  Could not query WSL status." -ForegroundColor DarkGray
    }
} else {
    Write-Host "Skipping WSL distro removal (-KeepWSL specified)." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Green
Write-Host "  Hermes Agent has been uninstalled."                 -ForegroundColor Green
Write-Host "=====================================================" -ForegroundColor Green
Write-Host ""
Read-Host "Press ENTER to exit"
