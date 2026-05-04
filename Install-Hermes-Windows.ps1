#Requires -Version 5.1
#Requires -RunAsAdministrator

# ==========================================
# Hermes Agent One-Click Windows Installer
# Installs: WSL2, Ubuntu-22.04, Ollama, Hermes Agent
# Creates: Desktop shortcuts + .bat launchers
# ==========================================

param(
    [switch]$Resume
)

$ErrorActionPreference = "Stop"
$ProgressPreference   = "Continue"

# --- Configuration ---
$DistroName     = "Ubuntu-22.04"
$InstallDir     = "$env:LOCALAPPDATA\HermesAgent"
$LaunchersDir   = "$InstallDir\Launchers"
$LogFile        = "$InstallDir\setup.log"
$ResumeFlag     = "$InstallDir\resume.flag"
$RepoName       = "NousResearch/hermes-agent"

# --- Logging ---
function Write-Log {
    param([string]$Message)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$ts | $Message"
    if (-not (Test-Path $InstallDir)) { New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null }
    $line | Tee-Object -FilePath $LogFile -Append | Write-Host
}

function Register-Resume {
    New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
    New-Item -ItemType File -Force -Path $ResumeFlag | Out-Null
    $scriptPath = $PSCommandPath
    $cmd = "powershell.exe -ExecutionPolicy Bypass -File `"$scriptPath`" -Resume"
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "HermesAgentSetup" -Value $cmd -ErrorAction SilentlyContinue
    Write-Log "Resume registered in RunOnce."
}

function Remove-Resume {
    Remove-Item -Path $ResumeFlag -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "HermesAgentSetup" -ErrorAction SilentlyContinue
}

function Test-WslInstalled {
    try {
        $null = wsl --status 2>$null
        return $LASTEXITCODE -eq 0
    } catch { return $false }
}

function Test-DistroInstalled {
    param([string]$Name)
    try {
        $list = wsl -l -v 2>$null
        return ($list -match [regex]::Escape($Name))
    } catch { return $false }
}

function Wait-DistroReady {
    param([string]$Name)
    Write-Log "Waiting for $Name to be ready..."
    for ($i = 0; $i -lt 30; $i++) {
        $out = wsl -d $Name -e true 2>&1
        if ($LASTEXITCODE -eq 0) { Write-Log "$Name is ready."; return $true }
        Start-Sleep -Seconds 2
    }
    return $false
}

# --- Banner ---
Write-Host ""
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "  Hermes Agent - One-Click Windows Installer"         -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host ""

if (-not $Resume) {
    Write-Log "=== Hermes Agent Windows Setup Started ==="
} else {
    Write-Log "=== Resuming Hermes Agent Setup ==="
}

# 1. Ensure WSL is installed
if (-not (Test-WslInstalled)) {
    Write-Log "WSL not detected. Installing WSL2 + $DistroName..."
    Write-Host "This enables the Virtual Machine Platform and may require a REBOOT." -ForegroundColor Yellow
    wsl --install -d $DistroName --no-launch
    Write-Log "WSL installation initiated. Registering resume..."
    Register-Resume
    Write-Host ""
    Write-Host "=====================================================" -ForegroundColor Green
    Write-Host "  REBOOT REQUIRED to finish enabling WSL2."         -ForegroundColor Green
    Write-Host "  After login, setup will resume automatically."    -ForegroundColor Green
    Write-Host "=====================================================" -ForegroundColor Green
    Write-Host ""
    Read-Host "Press ENTER to reboot now"
    Restart-Computer -Force
    exit 0
}

Write-Log "WSL is installed."

# 2. Ensure Ubuntu-22.04 is present
if (-not (Test-DistroInstalled -Name $DistroName)) {
    Write-Log "$DistroName not found. Installing..."
    wsl --install -d $DistroName --no-launch
} else {
    Write-Log "$DistroName is already installed."
}

# Set defaults
wsl --set-default-version 2 | Out-Null
wsl --set-default $DistroName | Out-Null
wsl --update | Out-Null

# 3. Wait for distro readiness
if (-not (Wait-DistroReady -Name $DistroName)) {
    Write-Log "ERROR: $DistroName did not become ready. Reboot and run again."
    exit 1
}

# 4. Update packages
Write-Log "Updating Ubuntu packages..."
wsl -d $DistroName -e bash -c "export DEBIAN_FRONTEND=noninteractive; sudo apt-get update -qq && sudo apt-get upgrade -y -qq" 2>&1 | ForEach-Object { Write-Log $_ }

Write-Log "Installing core dependencies inside WSL (curl, git, build-essential)..."
wsl -d $DistroName -e bash -c "export DEBIAN_FRONTEND=noninteractive; sudo apt-get install -y -qq curl git build-essential" 2>&1 | ForEach-Object { Write-Log $_ }

# 5. Install Ollama
Write-Log "Installing Ollama inside WSL..."
wsl -d $DistroName -e bash -c "curl -fsSL https://ollama.com/install.sh | sh" 2>&1 | ForEach-Object { Write-Log $_ }

# 6. Install Hermes Agent
Write-Log "Installing Hermes Agent inside WSL..."
wsl -d $DistroName -e bash -c "curl -fsSL https://raw.githubusercontent.com/$RepoName/main/scripts/install.sh | bash -s -- --skip-setup" 2>&1 | ForEach-Object { Write-Log $_ }

# 7. Pull a default small model so Ollama is useful out of the box
Write-Log "Pulling default lightweight model (llama3.2:3b) for Ollama..."
wsl -d $DistroName -e bash -c "ollama pull llama3.2:3b" 2>&1 | ForEach-Object { Write-Log $_ }

# 8. Create Windows .bat launchers
Write-Log "Creating launchers in $LaunchersDir ..."
New-Item -ItemType Directory -Force -Path $LaunchersDir | Out-Null

$hermesCli = @'
@echo off
cls
echo =====================================================
echo  Hermes Agent CLI
echo =====================================================
echo This window runs Hermes inside WSL (Ubuntu-22.04).
echo To quit, type /exit inside Hermes or close this window.
echo.
wsl -d Ubuntu-22.04 -e bash -lc "hermes chat"
echo.
echo Hermes CLI closed.
pause
'@
Set-Content -Path "$LaunchersDir\Hermes-CLI.bat" -Value $hermesCli

$hermesGateway = @'
@echo off
cls
echo =====================================================
echo  Hermes Gateway (Messaging Server)
echo =====================================================
echo Starting in background window...
start "Hermes Gateway" wsl -d Ubuntu-22.04 -e bash -lc "hermes gateway"
echo Gateway launched. Check the other window for output.
timeout /t 3 >nul
'@
Set-Content -Path "$LaunchersDir\Hermes-Gateway.bat" -Value $hermesGateway

$hermesAcp = @'
@echo off
cls
echo =====================================================
echo  Hermes ACP Server
echo =====================================================
echo Starting in background window...
start "Hermes ACP Server" wsl -d Ubuntu-22.04 -e bash -lc "hermes acp"
echo ACP Server launched. Check the other window for output.
timeout /t 3 >nul
'@
Set-Content -Path "$LaunchersDir\Hermes-ACP-Server.bat" -Value $hermesAcp

$ollamaServer = @'
@echo off
cls
echo =====================================================
echo  Ollama Server
echo =====================================================
echo Starting in background window...
start "Ollama Server" wsl -d Ubuntu-22.04 -e bash -lc "ollama serve"
echo Ollama server launched. Check the other window for output.
timeout /t 3 >nul
'@
Set-Content -Path "$LaunchersDir\Ollama-Server.bat" -Value $ollamaServer

$hermesSetup = @'
@echo off
cls
echo =====================================================
echo  Hermes Agent Interactive Setup
echo =====================================================
echo This runs the first-time configuration wizard.
echo.
wsl -d Ubuntu-22.04 -e bash -lc "hermes setup"
echo.
pause
'@
Set-Content -Path "$LaunchersDir\Hermes-Setup.bat" -Value $hermesSetup

# 9. Create Desktop shortcuts
$WshShell = New-Object -ComObject WScript.Shell
$Desktop  = [Environment]::GetFolderPath("Desktop")

$sc1 = $WshShell.CreateShortcut("$Desktop\Hermes CLI.lnk")
$sc1.TargetPath       = "$LaunchersDir\Hermes-CLI.bat"
$sc1.WorkingDirectory = $LaunchersDir
$sc1.IconLocation     = "cmd.exe,0"
$sc1.Description      = "Launch Hermes Agent interactive CLI"
$sc1.Save()

$sc2 = $WshShell.CreateShortcut("$Desktop\Hermes Gateway.lnk")
$sc2.TargetPath       = "$LaunchersDir\Hermes-Gateway.bat"
$sc2.WorkingDirectory = $LaunchersDir
$sc2.IconLocation     = "cmd.exe,0"
$sc2.Description      = "Launch Hermes Agent messaging gateway"
$sc2.Save()

$sc3 = $WshShell.CreateShortcut("$Desktop\Ollama Server.lnk")
$sc3.TargetPath       = "$LaunchersDir\Ollama-Server.bat"
$sc3.WorkingDirectory = $LaunchersDir
$sc3.IconLocation     = "cmd.exe,0"
$sc3.Description      = "Launch Ollama AI model server"
$sc3.Save()

$sc4 = $WshShell.CreateShortcut("$Desktop\Hermes Setup.lnk")
$sc4.TargetPath       = "$LaunchersDir\Hermes-Setup.bat"
$sc4.WorkingDirectory = $LaunchersDir
$sc4.IconLocation     = "cmd.exe,0"
$sc4.Description      = "Run Hermes Agent first-time setup"
$sc4.Save()

Write-Log "Desktop shortcuts created."

# 10. Create a Start Menu folder
$StartMenu = [Environment]::GetFolderPath("StartMenu")
$AppFolder = "$StartMenu\Programs\Hermes Agent"
New-Item -ItemType Directory -Force -Path $AppFolder | Out-Null

($WshShell.CreateShortcut("$AppFolder\Hermes CLI.lnk")).TargetPath       = "$LaunchersDir\Hermes-CLI.bat"; ($WshShell.CreateShortcut("$AppFolder\Hermes CLI.lnk")).WorkingDirectory = $LaunchersDir; ($WshShell.CreateShortcut("$AppFolder\Hermes CLI.lnk")).Save()
($WshShell.CreateShortcut("$AppFolder\Hermes Gateway.lnk")).TargetPath    = "$LaunchersDir\Hermes-Gateway.bat"; ($WshShell.CreateShortcut("$AppFolder\Hermes Gateway.lnk")).WorkingDirectory = $LaunchersDir; ($WshShell.CreateShortcut("$AppFolder\Hermes Gateway.lnk")).Save()
($WshShell.CreateShortcut("$AppFolder\Ollama Server.lnk")).TargetPath    = "$LaunchersDir\Ollama-Server.bat"; ($WshShell.CreateShortcut("$AppFolder\Ollama Server.lnk")).WorkingDirectory = $LaunchersDir; ($WshShell.CreateShortcut("$AppFolder\Ollama Server.lnk")).Save()
($WshShell.CreateShortcut("$AppFolder\Hermes Setup.lnk")).TargetPath     = "$LaunchersDir\Hermes-Setup.bat"; ($WshShell.CreateShortcut("$AppFolder\Hermes Setup.lnk")).WorkingDirectory = $LaunchersDir; ($WshShell.CreateShortcut("$AppFolder\Hermes Setup.lnk")).Save()
($WshShell.CreateShortcut("$AppFolder\Hermes ACP Server.lnk")).TargetPath = "$LaunchersDir\Hermes-ACP-Server.bat"; ($WshShell.CreateShortcut("$AppFolder\Hermes ACP Server.lnk")).WorkingDirectory = $LaunchersDir; ($WshShell.CreateShortcut("$AppFolder\Hermes ACP Server.lnk")).Save()

Write-Log "Start Menu folder created: Hermes Agent"

# 11. Cleanup
Remove-Resume

Write-Log "=== Setup Complete ==="
Write-Host ""
Write-Host "=====================================================" -ForegroundColor Green
Write-Host "  Hermes Agent is ready!"                           -ForegroundColor Green
Write-Host "=====================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Run 'Hermes Setup' on your desktop to configure providers." -ForegroundColor White
Write-Host "  2. Run 'Ollama Server' to start the local LLM backend." -ForegroundColor White
Write-Host "  3. Run 'Hermes CLI' to chat with your agent." -ForegroundColor White
Write-Host "  4. Run 'Hermes Gateway' to start the messaging server." -ForegroundColor White
Write-Host ""
Write-Host "Launchers folder: $LaunchersDir" -ForegroundColor DarkGray
Write-Host "Log file:        $LogFile" -ForegroundColor DarkGray
Write-Host ""
Read-Host "Press ENTER to exit setup"
