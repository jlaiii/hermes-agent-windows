@echo off
:: =====================================================
:: Hermes Agent Windows Setup Launcher
:: Double-click this file to run the installer.
:: It will auto-elevate to Administrator.
:: =====================================================
echo =====================================================
echo  Hermes Agent Setup Launcher
echo =====================================================
echo Detecting PowerShell...
where powershell >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: PowerShell not found on this system.
    echo Please install PowerShell 5.1 or later.
    pause
    exit /b 1
)

echo Starting setup as Administrator...
echo (You will see a UAC prompt. Click Yes to continue.)
echo.
powershell -ExecutionPolicy Bypass -Command "Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -File \"%~dp0Install-Hermes-Windows.ps1\"' -Verb RunAs -Wait"
if %errorlevel% neq 0 (
    echo.
    echo Setup exited with code %errorlevel%.
    pause
)
