@echo off
cd /d "%~dp0"

echo [BUILD] Closing any running PDFtoWord instances...
powershell -NoProfile -Command "Get-Process PDFtoWord -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue"

echo [BUILD] Closing Explorer windows that have dist\PDFtoWord open...
powershell -NoProfile -Command "$sh = New-Object -ComObject Shell.Application; $sh.Windows() | Where-Object { try { $_.LocationURL -like '*PDFtoWord*' } catch { $false } } | ForEach-Object { $_.Quit() }; Start-Sleep -Milliseconds 800"

echo [BUILD] Cleaning old artifacts...
if exist "%CD%\dist\PDFtoWord" (
    rd /s /q "%CD%\dist\PDFtoWord" 2>nul
)
if exist "%CD%\dist\PDFtoWord" (
    echo [ERROR] dist\PDFtoWord is still locked. Close any Explorer window showing that folder,
    echo         then run the build again.
    pause
    exit /b 1
)
if exist "%CD%\dist\PDFtoWord.zip" (
    del /f /q "%CD%\dist\PDFtoWord.zip" 2>nul
)

echo [BUILD] Running PyInstaller...
venv\Scripts\pyinstaller.exe pdf2word.spec --noconfirm
if errorlevel 1 (
    echo [ERROR] PyInstaller failed.
    pause
    exit /b 1
)

echo [BUILD] Waiting for file locks to release...
timeout /t 5 /nobreak > nul

echo [BUILD] Creating zip...
powershell -NoProfile -Command ^
    "try { Compress-Archive -Path '%CD%\dist\PDFtoWord' -DestinationPath '%CD%\dist\PDFtoWord.zip' -Force; exit 0 } catch { Write-Error $_; exit 1 }"
if errorlevel 1 (
    echo [WARNING] ZIP creation failed. dist\PDFtoWord\ is ready but PDFtoWord.zip was not created.
    echo           You can zip dist\PDFtoWord\ manually.
) else (
    echo [BUILD] Done. dist\PDFtoWord\ and dist\PDFtoWord.zip are ready.
)
pause
