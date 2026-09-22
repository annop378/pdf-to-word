@echo off
echo Restarting Windows Explorer to release file locks...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 2 /nobreak >nul
start explorer.exe
echo Done. Explorer restarted.
