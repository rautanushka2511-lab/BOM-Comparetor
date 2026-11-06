#!/usr/bin/env python3
import os
import subprocess
import shutil
from pathlib import Path

# Use current working directory
print(f"Working directory: {os.getcwd()}\n")

# Step 1: Remove directories
dirs_to_remove = [
    "final_app",
    "build", 
    "dist",
    r"Deployment\win-unpacked",
    r"Deployment\TCUA_BOM_App",
    r"Deployment\TCUA BOM Compare_v1.0.0"
]

print("=" * 60)
print("STEP 1: REMOVING DIRECTORIES")
print("=" * 60)

for dir_path in dirs_to_remove:
    if os.path.exists(dir_path):
        print(f"Removing: {dir_path}")
        shutil.rmtree(dir_path)
        print(f"  [OK] Removed")
    else:
        print(f"[SKIP] {dir_path} not found")

# Step 2: Run npm dist
print("\n" + "=" * 60)
print("STEP 2: RUNNING NPM BUILD")
print("=" * 60 + "\n")

result = subprocess.run("npm run dist", shell=True)
if result.returncode == 0:
    print("\n[OK] npm run dist completed successfully")
else:
    print(f"\n[ERROR] npm run dist failed with code {result.returncode}")

# Step 3: Find all exe files
print("\n" + "=" * 60)
print("STEP 3: FINDING ALL EXE FILES IN DEPLOYMENT")
print("=" * 60 + "\n")

exe_files = []
deployment_path = Path("Deployment")

if deployment_path.exists():
    for exe_path in deployment_path.rglob("*.exe"):
        exe_files.append(str(exe_path))
        print(f"[FOUND] {exe_path}")

print(f"\nTotal EXE files found: {len(exe_files)}")

# Step 4: Organize deployment
print("\n" + "=" * 60)
print("STEP 4: ORGANIZING DEPLOYMENT - KEEPING ONE EXE")
print("=" * 60 + "\n")

if len(exe_files) == 0:
    print("[INFO] No EXE files found")
elif len(exe_files) == 1:
    print(f"[INFO] Only one EXE found: {exe_files[0]}")
else:
    print(f"[INFO] Found {len(exe_files)} EXE files")
    
    # Find top-level exe (directly in Deployment folder)
    top_level_exe = None
    nested_exes = []
    
    for exe_path in exe_files:
        # Normalize path separators
        normalized = exe_path.replace("/", "\\")
        # Check if it's in Deployment root
        parts = Path(normalized).parts
        if len(parts) == 2 and parts[0] == "Deployment":
            top_level_exe = exe_path
        else:
            nested_exes.append(exe_path)
    
    if top_level_exe:
        print(f"[DECISION] Keeping top-level EXE: {top_level_exe}")
        print("[ACTION] Removing directories with nested EXEs")
        
        # Remove all subdirectories containing exe files
        dirs_removed = set()
        for nested_exe in nested_exes:
            nested_path = Path(nested_exe)
            parent_dir = nested_path.parent
            
            # Find the subdirectory to remove (the first subdir of Deployment)
            if parent_dir.name != "Deployment":
                parts = parent_dir.parts
                for i, part in enumerate(parts):
                    if part == "Deployment" and i + 1 < len(parts):
                        subdir = Path(*parts[:i+2])
                        if str(subdir) not in dirs_removed:
                            dirs_removed.add(str(subdir))
                            print(f"  Removing: {subdir}")
                            if subdir.exists():
                                shutil.rmtree(str(subdir))
    else:
        print(f"[INFO] No top-level EXE found, keeping first nested: {nested_exes[0] if nested_exes else 'None'}")

# Step 5: Final list
print("\n" + "=" * 60)
print("STEP 5: FINAL LIST OF ALL EXE FILES IN DEPLOYMENT")
print("=" * 60 + "\n")

final_exes = []
if deployment_path.exists():
    for exe_path in deployment_path.rglob("*.exe"):
        final_exes.append(str(exe_path))
        print(f"  {exe_path}")

if not final_exes:
    print("  [No EXE files found]")
else:
    print(f"\nTotal EXE files remaining: {len(final_exes)}")

print("\n" + "=" * 60)
print("COMPLETE")
print("=" * 60)
