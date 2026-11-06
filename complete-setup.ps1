# Automated setup script for TCUA BOM Comparison deployment
# Run this script to complete the backend setup

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "TCUA BOM Comparison - Deployment Setup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$appDir = "D:\akshar\apps\TCUA Perfect"
$backendDir = "$appDir\Deployment\win-unpacked\resources\backend"
$venvDir = "$appDir\.venv"

Write-Host "`n[1/4] Verifying directories..." -ForegroundColor Yellow
if (-not (Test-Path $backendDir)) {
    Write-Host "ERROR: Backend directory not found: $backendDir" -ForegroundColor Red
    exit 1
}
Write-Host "✓ Backend directory exists" -ForegroundColor Green

Write-Host "`n[2/4] Checking for Python in venv..." -ForegroundColor Yellow
$pythonExe = "$venvDir\Scripts\python.exe"
if (Test-Path $pythonExe) {
    Write-Host "✓ Found Python in venv: $pythonExe" -ForegroundColor Green

    Write-Host "`n[3/4] Copying Python environment..." -ForegroundColor Yellow
    try {
        # Copy Python executable
        Copy-Item -Path $pythonExe -Destination $backendDir -Force
        Write-Host "✓ Copied python.exe" -ForegroundColor Green

        # Copy site-packages (contains all dependencies)
        $sitePackages = "$venvDir\Lib\site-packages"
        if (Test-Path $sitePackages) {
            Write-Host "Copying site-packages (this may take a while)..." -ForegroundColor Yellow
            Copy-Item -Path $sitePackages -Destination "$backendDir\site-packages" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "✓ Copied site-packages" -ForegroundColor Green
        }

        # Copy Lib folder (Python standard library)
        $libDir = "$venvDir\Lib"
        if (Test-Path "$libDir\*.pyc") {
            Write-Host "Copying Python standard library..." -ForegroundColor Yellow
            Copy-Item -Path "$libDir" -Destination "$backendDir\Lib" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "✓ Copied Lib directory" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "WARNING: Failed to copy some Python files: $_" -ForegroundColor Yellow
        Write-Host "The application may still work if system Python is available" -ForegroundColor Yellow
    }
} else {
    Write-Host "⚠ Python not found in venv" -ForegroundColor Yellow
    Write-Host "Checking for system Python..." -ForegroundColor Yellow

    $systemPython = (Get-Command python.exe -ErrorAction SilentlyContinue).Source
    if ($systemPython) {
        Write-Host "✓ Found system Python: $systemPython" -ForegroundColor Green
        Write-Host "The batch file will use system Python" -ForegroundColor Yellow
    } else {
        Write-Host "ERROR: No Python found in venv or system!" -ForegroundColor Red
        Write-Host "Please install Python before proceeding" -ForegroundColor Red
        exit 1
    }
}

Write-Host "`n[4/4] Verifying backend files..." -ForegroundColor Yellow
$requiredFiles = @(
    "app.py",
    "bom_comp_levelcs.py",
    "generate_json_hierarchy.py",
    "config91112.json",
    "submodule_structure.json",
    "app.exe.bat"
)

$missingFiles = @()
foreach ($file in $requiredFiles) {
    $filePath = "$backendDir\$file"
    if (Test-Path $filePath) {
        Write-Host "✓ $file" -ForegroundColor Green
    } else {
        Write-Host "✗ MISSING: $file" -ForegroundColor Red
        $missingFiles += $file
    }
}

if ($missingFiles.Count -gt 0) {
    Write-Host "`nERROR: Missing required files: $($missingFiles -join ', ')" -ForegroundColor Red
    exit 1
}

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "✓ Setup Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green

Write-Host "`nYou can now run the application:" -ForegroundColor Cyan
Write-Host "  $appDir\Deployment\win-unpacked\TCUA BOM Compare.exe" -ForegroundColor Cyan

Write-Host "`nFiles will be deployed to:" -ForegroundColor Cyan
Write-Host "  $backendDir\" -ForegroundColor Cyan

Write-Host "`nBackend directory contents:" -ForegroundColor Yellow
Get-ChildItem -Path $backendDir -Force | ForEach-Object {
    $size = if ($_.PSIsContainer) { "[DIR]" } else { "[$('{0:F1}' -f ($_.Length/1MB)) MB]" }
    Write-Host "  $($_.Name) $size" -ForegroundColor Gray
}

Write-Host "`n Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

