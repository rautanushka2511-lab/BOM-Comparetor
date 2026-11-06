# TCUA BOM Comparison - Quick Start Guide

## For End Users - Running the Application

### Option 1: Simple Run (Recommended if Python is Installed)

1. Open the Deployment folder:
   ```
   Deployment\win-unpacked\
   ```

2. Double-click:
   ```
   TCUA BOM Compare.exe
   ```

3. Wait for the application to start (first run may take a few seconds)

### Option 2: Setup & Run (If Python Not Available)

1. Run the setup script:
   ```powershell
   # Open PowerShell as Administrator in the project folder
   Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
  ./complete-setup.ps1
   ```

2. After setup completes, run:
   ```
   Deployment\win-unpacked\TCUA BOM Compare.exe
   ```

---

## For IT Admins - Deployment

### Requirements
- Windows 10/11 (64-bit)
- Either:
  - Python 3.9+ installed and in PATH, OR
  - The complete `Deployment\win-unpacked\` folder with Python bundled

### Deployment Steps

1. **Extract Files**
   - Extract the entire `Deployment\win-unpacked\` folder to desired location
   - Example: `C:\Program Files\TCUA BOM Comparison\`

2. **Optional: Add to PATH** (if not using bundled Python)
   - Ensure Python is installed system-wide
   - Or run `complete-setup.ps1` to bundle Python

3. **Create Shortcut** (Optional)
   - Create a desktop shortcut pointing to:
     ```
     Target: C:\Program Files\TCUA BOM Comparison\TCUA BOM Compare.exe
     ```

4. **First Run**
   - Double-click the EXE
   - Application will launch Electron and Python backend
   - UI will load at http://127.0.0.1:9112

### Troubleshooting

**Issue**: "Python not found"
- **Solution**: Install Python from python.org or run `complete-setup.ps1`

**Issue**: Port 9112 already in use
- **Solution**: Edit `resources/backend/app.exe.bat` and change PORT value
  ```batch
  set "PORT=9113"
  ```

**Issue**: Application doesn't start
- **Check**: Look for error console window (may appear briefly)
- **Debug**: Run this in PowerShell from backend folder:
  ```powershell
  cd Deployment\win-unpacked\resources\backend
  python app.py
  ```

---

## System Requirements

| Component | Requirement |
|-----------|-------------|
| OS | Windows 10/11 (64-bit) |
| Memory | 4 GB RAM minimum, 8 GB recommended |
| Disk Space | 1-3 GB (depends on Python libraries) |
| Python | 3.9, 3.10, 3.11, or 3.12 |
| Disk Speed | SSD recommended for first startup |

---

## Feature Overview

### Main Functions
1. **Upload Excel Files**: Compare two BOM files
2. **Dynamic Column Mapping**: Map custom column names
3. **Comparison Results**:
   - Excel report with detailed comparison
   - JSON hierarchy view
   - Summary with statistics

### Supported File Types
- `.xlsx` (Excel 2007+)
- `.xls` (Excel 97-2003)

### Configuration
- `config91112.json`: Segment rules and comparison thresholds
- `submodule_structure.json`: Submodule location mapping

---

## Support

For issues or questions:
1. Check `DEPLOYMENT_GUIDE.md` for detailed setup instructions
2. Contact the development team

---

**Version**: 1.0.0
**Last Updated**: June 2, 2026

