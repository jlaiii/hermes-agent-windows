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
