@echo off
cd /d "%~dp0"

echo [BUILD] Cleaning old artifacts...
powershell -NoProfile -Command "if (Test-Path '%CD%\dist\PDFtoWord') { Remove-Item -Recurse -Force '%CD%\dist\PDFtoWord' }; if (Test-Path '%CD%\dist\PDFtoWord.zip') { Remove-Item -Force '%CD%\dist\PDFtoWord.zip' }"

echo [BUILD] Running PyInstaller...
venv\Scripts\pyinstaller.exe pdf2word.spec --noconfirm
if errorlevel 1 (
    echo [ERROR] PyInstaller failed.
    pause
    exit /b 1
)

echo [BUILD] Creating zip...
powershell -NoProfile -Command "Compress-Archive -Path '%CD%\dist\PDFtoWord' -DestinationPath '%CD%\dist\PDFtoWord.zip' -Force"

echo [BUILD] Done. dist\PDFtoWord\ and dist\PDFtoWord.zip are ready.
pause
