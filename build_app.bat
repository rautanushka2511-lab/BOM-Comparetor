@echo off
cd /d "C:\Users\GSS527189\Documents\BCD_akshar\apps\TCUA Perfect"
echo === Step 1: Removing directories ===
for %%D in (final_app build dist) do (
    if exist "%%D" (
        echo Removing %%D...
        rmdir /s /q "%%D" 2>nul
    )
)

if exist "Deployment\win-unpacked" rmdir /s /q "Deployment\win-unpacked" 2>nul
if exist "Deployment\TCUA_BOM_App" rmdir /s /q "Deployment\TCUA_BOM_App" 2>nul
if exist "Deployment\TCUA BOM Compare_v1.0.0" rmdir /s /q "Deployment\TCUA BOM Compare_v1.0.0" 2>nul

echo === Step 2: Running npm run dist ===
call npm run dist

echo === Step 3: Listing final EXE files ===
dir /s/b "Deployment\*.exe"

pause
