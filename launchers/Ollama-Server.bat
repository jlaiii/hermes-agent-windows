@echo off
cls
echo =====================================================
echo  Ollama Server
echo =====================================================
echo Starting in background window...
start "Ollama Server" wsl -d Ubuntu-22.04 -e bash -lc "ollama serve"
echo Ollama server launched. Check the other window for output.
timeout /t 3 >nul
