@echo off
cls
echo =====================================================
echo  Hermes ACP Server
echo =====================================================
echo Starting in background window...
start "Hermes ACP Server" wsl -d Ubuntu-22.04 -e bash -lc "hermes acp"
echo ACP Server launched. Check the other window for output.
timeout /t 3 >nul
