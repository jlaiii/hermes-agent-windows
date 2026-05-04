@echo off
cls
echo =====================================================
echo  Hermes Gateway (Messaging Server)
echo =====================================================
echo Starting in background window...
start "Hermes Gateway" wsl -d Ubuntu-22.04 -e bash -lc "hermes gateway"
echo Gateway launched. Check the other window for output.
timeout /t 3 >nul
