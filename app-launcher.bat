@echo off
REM TCUA BOM Comparison - Backend Launcher
REM This batch file launches the FastAPI backend using the virtual environment Python

setlocal enabledelayedexpansion

REM Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"
cd /d "!SCRIPT_DIR!"

REM Check if we're in the bundled environment or development
if exist ".venv" (
    REM Development environment
    set "PYTHON_EXE=.venv\Scripts\python.exe"
    set "APP_DIR=."
) else if exist "..\.venv" (
    REM From bundled resources
    set "PYTHON_EXE=..\.venv\Scripts\python.exe"
    set "APP_DIR=."
) else (
    REM Try system Python
    set "PYTHON_EXE=python"
    set "APP_DIR=."
)

REM Set port if not already set
if not defined PORT (
    set "PORT=9112"
)

REM Run the application
"!PYTHON_EXE!" app.py --port !PORT!

