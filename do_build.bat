@echo off
setlocal enabledelayedexpansion

cd /d "C:\Users\GSS527189\Documents\BCD_akshar\apps\TCUA Perfect"

echo.
echo === STEP 1: REMOVING DIRECTORIES ===
echo.

if exist "final_app" (
    echo Removing final_app...
    rmdir /s /q "final_app"
    echo [OK] final_app removed
) else (
    echo [SKIP] final_app not found
)

if exist "build" (
    echo Removing build...
    rmdir /s /q "build"
    echo [OK] build removed
) else (
    echo [SKIP] build not found
)

if exist "dist" (
    echo Removing dist...
    rmdir /s /q "dist"
    echo [OK] dist removed
) else (
    echo [SKIP] dist not found
)

if exist "Deployment\win-unpacked" (
    echo Removing Deployment\win-unpacked...
    rmdir /s /q "Deployment\win-unpacked"
    echo [OK] Deployment\win-unpacked removed
) else (
    echo [SKIP] Deployment\win-unpacked not found
)

if exist "Deployment\TCUA_BOM_App" (
    echo Removing Deployment\TCUA_BOM_App...
    rmdir /s /q "Deployment\TCUA_BOM_App"
    echo [OK] Deployment\TCUA_BOM_App removed
) else (
    echo [SKIP] Deployment\TCUA_BOM_App not found
)

if exist "Deployment\TCUA BOM Compare_v1.0.0" (
    echo Removing Deployment\TCUA BOM Compare_v1.0.0...
    rmdir /s /q "Deployment\TCUA BOM Compare_v1.0.0"
    echo [OK] Deployment\TCUA BOM Compare_v1.0.0 removed
) else (
    echo [SKIP] Deployment\TCUA BOM Compare_v1.0.0 not found
)

echo.
echo === STEP 2: RUNNING NPM BUILD ===
echo.

call npm run dist

echo.
echo === STEP 3: FINDING ALL EXE FILES IN DEPLOYMENT ===
echo.

setlocal enabledelayedexpansion
set "exe_count=0"
set "top_level_exe="
set "nested_exe="

for /r "Deployment" %%F in (*.exe) do (
    set /a exe_count+=1
    set "full_path=%%F"
    set "file_name=%%~nxF"
    echo [FOUND] !full_path!
    
    REM Check if it's in Deployment root (top-level)
    if "!full_path:Deployment\=!" equ "!file_name!" (
        if "!top_level_exe!" equ "" (
            set "top_level_exe=!full_path!"
            echo [TOP-LEVEL] !full_path!
        )
    ) else (
        if "!nested_exe!" equ "" (
            set "nested_exe=!full_path!"
            echo [NESTED] !full_path!
        )
    )
)

echo.
echo === STEP 4: ORGANIZING DEPLOYMENT - KEEPING ONE EXE ===
echo.

if !exe_count! equ 0 (
    echo [INFO] No EXE files found in Deployment
) else if !exe_count! equ 1 (
    echo [INFO] Only one EXE found, keeping as-is
) else (
    echo [INFO] Found !exe_count! EXE files
    
    if not "!top_level_exe!" equ "" (
        echo [DECISION] Keeping top-level EXE: !top_level_exe!
        echo [ACTION] Removing subdirectories with EXEs
        
        REM Remove all subdirectories in Deployment that contain exe files
        for /d %%D in (Deployment\*) do (
            if not "%%D" equ "Deployment" (
                set "subdir_has_exe=0"
                for /r "%%D" %%F in (*.exe) do (
                    set "subdir_has_exe=1"
                )
                if !subdir_has_exe! equ 1 (
                    echo Removing directory: %%D
                    rmdir /s /q "%%D"
                )
            )
        )
    ) else if not "!nested_exe!" equ "" (
        echo [DECISION] Keeping nested EXE: !nested_exe!
    )
)

echo.
echo === STEP 5: FINAL LIST OF ALL EXE FILES IN DEPLOYMENT ===
echo.

set "final_count=0"
for /r "Deployment" %%F in (*.exe) do (
    set /a final_count+=1
    echo [!final_count!] %%F
)

if !final_count! equ 0 (
    echo No EXE files found in Deployment
) else (
    echo.
    echo Total EXE files: !final_count!
)

echo.
echo === COMPLETE ===
echo.
