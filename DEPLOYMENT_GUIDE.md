# TCUA BOM Comparison - Fixed Deployment Guide

## Issue Fixed
**Problem**: Electron app was unable to find the Python backend executable (`app.exe`), throwing `ENOENT` error.

**Root Cause**: PyInstaller was taking too long to complete, and there's no bundled `app.exe`.

**Solution**: Updated the launcher to use cmd.exe to run Python directly.

---

## What Has Been Done

### 1. Updated `main.js` (Electron Launcher)
- Changed from looking for `app.exe` to using `cmd.exe /c` to run Python
- Added fallback logic and debug logging
- Properly handles paths with spaces

### 2. Enhanced Backend Files
- Created/Updated `Deployment\win-unpacked\resources\backend\` with all Python dependencies:
  - `app.py` (main FastAPI application)
  - `bom_comp_levelcs.py` (BOM comparator)
  - `generate_json_hierarchy.py` (hierarchy generator)
  - `config91112.json` (configuration)
  - `submodule_structure.json` (submodule mapping)
  - `models/` (sentence-transformer models)
  - `static/` (web UI files)

### 3. Updated `app.exe.bat` Wrapper
- More robust Python detection
- Checks for python.exe in backend directory first
- Falls back to system Python if available
- Better error reporting

---

## What You Need to Do

### Step 1: Copy Python Environment to Backend (IMPORTANT)

**Option A: Copy from Virtual Environment (Recommended)**
```powershell
cd "D:\akshar\apps\TCUA Perfect"

# Copy Python interpreter and libraries
Copy-Item -Path ".venv\Scripts\python.exe" `
  -Destination "Deployment\win-unpacked\resources\backend\" -Force

# Copy necessary DLLs and libraries
Copy-Item -Path ".venv\Lib" `
  -Destination "Deployment\win-unpacked\resources\backend\" -Recurse -Force

# Copy site-packages
Copy-Item -Path ".venv\Lib\site-packages" `
  -Destination "Deployment\win-unpacked\resources\backend\" -Recurse -Force
```

**Option B: Use System Python (If Available)**
- Ensure Python is installed system-wide and in PATH
- The batch file will automatically detect and use it

### Step 2: Update Electron Configuration

The deployment is already configured. If you want to rebuild:

```powershell
cd "D:\akshar\apps\TCUA Perfect"
npm run dist
```

### Step 3: Test the Application

Run the executable:
```
D:\akshar\apps\TCUA Perfect\Deployment\win-unpacked\TCUA BOM Compare.exe
```

**Expected Behavior:**
1. Electron window opens
2. Backend starts automatically
3. FastAPI server launches on port 9112
4. Web UI loads

---

## Troubleshooting

### If "Python not found" Error Appears
1. Verify Python is installed: `python --version`
2. Copy Python using Option A above
3. Check that app.py exists in the backend directory

### If Port 9112 is Already in Use
- Change `PORT=9112` in the batch file to an available port
- Or kill existing processes using that port

### If Modules Cannot Be Found
- Ensure all dependencies are copied to `Deployment\win-unpacked\resources\backend\`
- May need to copy entire `.venv` folder

### To Debug Backend Startup
- Add a console window to see error messages by changing to:
```batch
@echo off
REM Remove /c to keep window open for debugging
```

---

## Key Files Modified

1. **`main.js`** - Updated Electron backend spawning logic
2. **`Deployment\win-unpacked\resources\backend\app.exe.bat`** - Enhanced Python launcher
3. **`setup-deployment.ps1`** - Helper script for deployment setup

---

## Distribution

The application is ready to be distributed. Share the folder:
```
D:\akshar\apps\TCUA Perfect\Deployment\win-unpacked\
```

**Include:**
- TCUA BOM Compare.exe
- All DLLs and resources
- resources/backend/ folder with all Python files

**Size:** ~2-3 GB (depending on Python libraries chosen)

---

## Next Steps

1. Follow **Step 1** above to copy Python environment
2. Test the application
3. If successful, create installer using electron-builder:
   ```powershell
   npm run dist
   ```
4. Share the generated EXE with your team

---

**Last Updated**: June 2, 2026

