@echo off
chcp 65001 >nul
echo ========================================
echo  PDF to Word - 打包 EXE
echo ========================================

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] 找不到 Python，請確認 PATH 設定
    pause & exit /b 1
)

echo [1/4] 安裝相依套件...
python -m pip install pyinstaller --quiet
if errorlevel 1 ( echo [ERROR] pip install 失敗 & pause & exit /b 1 )

echo [2/4] 清除舊的 build...
if exist dist\PdfToWord rmdir /s /q dist\PdfToWord
if exist build rmdir /s /q build

echo [3/4] 打包中（使用現有 spec）...
python -m PyInstaller pdf2word.spec
if errorlevel 1 ( echo [ERROR] PyInstaller 打包失敗 & pause & exit /b 1 )

:: spec 輸出名稱為 PDF轉Word，重新命名為 PdfToWord（Launcher 路徑一致、避免中文路徑問題）
if exist "dist\PDF轉Word" (
    if exist "dist\PdfToWord" rmdir /s /q "dist\PdfToWord"
    move "dist\PDF轉Word" "dist\PdfToWord" >nul
)

echo [4/4] 打包 zip...
if exist dist\PdfToWord.zip del /q dist\PdfToWord.zip
powershell -Command "Compress-Archive -Path 'dist\PdfToWord' -DestinationPath 'dist\PdfToWord.zip' -Force"
if errorlevel 1 ( echo [ERROR] 打 zip 失敗 & pause & exit /b 1 )

echo.
echo ========================================
echo  完成！
echo  EXE 資料夾：dist\PdfToWord\
echo  ZIP 套件：  dist\PdfToWord.zip  ^<-- 複製到網路磁碟
echo ========================================
start "" "%~dp0dist"
pause
