@echo off
REM CAN Test Tool Launcher for Windows
REM =====================================

echo Starting CAN Test Tool - FSCM/RSCM
echo.

REM Check Python
py --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python not found. Please install Python 3.7+
    pause
    exit /b 1
)

REM Get script directory
set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

REM Install python-can if missing
py -m pip show python-can >nul 2>&1
if errorlevel 1 (
    echo Installing python-can library...
    py -m pip install python-can
)

echo.
echo Launching CAN Test Tool...
echo Supported hardware: PCAN-USB
echo Virtual mode available for testing without hardware
echo.
py can_test_tool.py

if errorlevel 1 (
    echo.
    echo Application exited with error. Press any key to close.
    pause >nul
)
