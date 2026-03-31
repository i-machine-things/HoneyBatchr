@echo off
REM Quick setup script for Honey Batchr developers
REM This installs dependencies and runs all necessary setup steps

setlocal enabledelayedexpansion

echo.
echo ========================================
echo Honey Batchr - Developer Setup
echo ========================================
echo.

REM Check Python installation
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found
    echo Please install Python 3.8+
    pause
    exit /b 1
)

echo Installing dependencies...
pip install -r requirements.txt --quiet

if errorlevel 1 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo Dependencies installed successfully!
echo.
echo Next steps:
echo 1. Run: build.bat (to build the executable)
echo 2. Run: register_context_menu.bat (as admin, to add context menu)
echo 3. Run: dist\HoneyBatchr\HoneyBatchr.exe (to launch the app)
echo.
pause
