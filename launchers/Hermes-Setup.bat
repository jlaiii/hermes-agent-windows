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
