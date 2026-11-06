@echo off
setlocal

:: 1) Stop running app/backend from previous unpacked build
echo Stopping existing processes...
taskkill /F /IM "TCUA BOM Compare.exe" /T >nul 2>&1
taskkill /F /IM "app.exe" /T >nul 2>&1

:: 2) Clean old outputs
echo Cleaning old outputs...
if exist ".\Deployment\win-unpacked" rd /s /q ".\Deployment\win-unpacked"
if exist ".\Deployment\TCUA_BOM_Compare_win-unpacked.zip" del /f /q ".\Deployment\TCUA_BOM_Compare_win-unpacked.zip"

:: 3) Rebuild Python backend
echo Rebuilding Python backend...
if exist "dist\app\app.exe" (
    echo [INFO] dist\app\app.exe already exists. Skipping PyInstaller compilation to save time.
    echo [INFO] Delete the "dist" folder if you want to force a full Python rebuild.
) else (
    call npm run build:py
)

:: 4) Build unpacked Electron app without signing
echo Building unpacked Electron app...
set CSC_IDENTITY_AUTO_DISCOVERY=false
call npx electron-builder --win --x64 --dir --config.win.sign=false --config.win.signAndEditExecutable=false --publish never

:: 5) Zip for sharing
echo Zipping for sharing...
tar.exe -acf .\Deployment\TCUA_BOM_Compare_win-unpacked.zip -C .\Deployment\win-unpacked .

echo Build Complete!
