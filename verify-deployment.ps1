#!/usr/bin/env powershell
# Verification and troubleshooting script

Write-Host "TCUA BOM Comparison - Verification Script" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

$appDir = "D:\akshar\apps\TCUA Perfect"
$deployDir = "$appDir\Deployment\win-unpacked"
$backendDir = "$deployDir\resources\backend"

Write-Host "`n[Checking Application Structure]" -ForegroundColor Yellow

# Check main EXE
if (Test-Path "$deployDir\TCUA BOM Compare.exe") {
    Write-Host "✓ Main executable found" -ForegroundColor Green
} else {
    Write-Host "✗ Main executable NOT found" -ForegroundColor Red
}

# Check main.js
if (Test-Path "$appDir\main.js") {
    Write-Host "✓ main.js exists" -ForegroundColor Green
    $content = Get-Content "$appDir\main.js" | Select-String "cmd.exe"
    if ($content) {
        Write-Host "✓ main.js updated to use cmd.exe" -ForegroundColor Green
    } else {
        Write-Host "⚠ main.js may not be updated correctly" -ForegroundColor Yellow
    }
} else {
    Write-Host "✗ main.js NOT found" -ForegroundColor Red
}

# Check backend files
Write-Host "`n[Checking Backend Software]" -ForegroundColor Yellow

$requiredFiles = @(
    "app.py",
    "bom_comp_levelcs.py",
    "generate_json_hierarchy.py",
    "config91112.json",
    "submodule_structure.json",
    "app.exe.bat"
)

foreach ($file in $requiredFiles) {
    $fullPath = "$backendDir\$file"
    if (Test-Path $fullPath) {
        $size = if ((Get-Item $fullPath).PSIsContainer) {
            "[DIR]"
        } else {
            "[$('{0:F1}' -f ((Get-Item $fullPath).Length/1KB)) KB]"
        }
        Write-Host "✓ $file $size" -ForegroundColor Green
    } else {
        Write-Host "✗ $file NOT found" -ForegroundColor Red
    }
}

# Check Python
Write-Host "`n[Checking Python Availability]" -ForegroundColor Yellow

$localPython = "$backendDir\python.exe"
$systemPython = Get-Command python.exe -ErrorAction SilentlyContinue

if (Test-Path $localPython) {
    Write-Host "✓ Local Python found in backend" -ForegroundColor Green
    & $localPython --version 2>&1 | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
} elseif ($systemPython) {
    Write-Host "✓ System Python available" -ForegroundColor Green
    & python --version 2>&1 | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
} else {
    Write-Host "✗ NO PYTHON FOUND" -ForegroundColor Red
    Write-Host "  Please install Python or run complete-setup.ps1" -ForegroundColor Yellow
}

# Check for required Python modules
if ($localPython -or $systemPython) {
    Write-Host "`n[Checking Python Dependencies]" -ForegroundColor Yellow

    $modules = @("fastapi", "uvicorn", "pandas", "numpy", "openpyxl")
    $pythonCmd = if (Test-Path $localPython) { $localPython } else { "python" }

    foreach ($module in $modules) {
        $result = & $pythonCmd -c "import $module; print('$module.__version__')" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ $module" -ForegroundColor Green
        } else {
            Write-Host "⚠ $module (may not be installed)" -ForegroundColor Yellow
        }
    }
}

# Summary
Write-Host "`n[Summary]" -ForegroundColor Yellow

$issues = @()
if (-not (Test-Path "$deployDir\TCUA BOM Compare.exe")) { $issues += "Main executable missing" }
if (-not (Test-Path "$backendDir\app.py")) { $issues += "Backend Python files missing" }
if (-not (Test-Path $localPython) -and -not $systemPython) { $issues += "Python not found" }

if ($issues.Count -eq 0) {
    Write-Host "`n✓ All checks passed! Application should be ready to run." -ForegroundColor Green
    Write-Host "`nYou can now run: $deployDir\TCUA BOM Compare.exe" -ForegroundColor Cyan
} else {
    Write-Host "`n⚠ Issues found:" -ForegroundColor Yellow
    $issues | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    Write-Host "`nTo fix: Run complete-setup.ps1" -ForegroundColor Cyan
}

Write-Host "`nPress any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

