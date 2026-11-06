#!/bin/bash
# PowerShell version for manual deployment setup

# Create the backend folder structure
New-Item -Path "Deployment\win-unpacked\resources\backend" -ItemType Directory -Force | Out-Null

# Copy app.py and all dependencies
Copy-Item -Path "app.py" -Destination "Deployment\win-unpacked\resources\backend\" -Force
Copy-Item -Path "bom_comp_levelcs.py" -Destination "Deployment\win-unpacked\resources\backend\" -Force
Copy-Item -Path "generate_json_hierarchy.py" -Destination "Deployment\win-unpacked\resources\backend\" -Force
Copy-Item -Path "config91112.json" -Destination "Deployment\win-unpacked\resources\backend\" -Force
Copy-Item -Path "submodule_structure.json" -Destination "Deployment\win-unpacked\resources\backend\" -Force
Copy-Item -Path "static" -Destination "Deployment\win-unpacked\resources\backend\" -Recurse -Force
Copy-Item -Path "models" -Destination "Deployment\win-unpacked\resources\backend\" -Recurse -Force

# Create a wrapper batch file that runs Python
$wrapper = @'
@echo off
cd %~dp0
REM Run the Python app
python app.py
'@

$wrapper | Out-File -FilePath "Deployment\win-unpacked\resources\backend\app.exe.bat" -Encoding ASCII

echo "Deployment setup complete!"

