@echo off
REM Honey Batchr Build Script

setlocal enabledelayedexpansion

echo.
echo ========================================
echo Honey Batchr Build Script
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.8+
    pause
    exit /b 1
)

REM Check if required packages are installed
echo Checking dependencies...
python -m pip show PyQt6 >nul 2>&1
if errorlevel 1 (
    echo Installing PyQt6...
    python -m pip install PyQt6 --quiet
)

python -m pip show PyInstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    python -m pip install PyInstaller --quiet
)

python -m pip show Pillow >nul 2>&1
if errorlevel 1 (
    echo Installing Pillow...
    python -m pip install Pillow --quiet
)

REM Create icons
echo.
echo Step 1: Generating icons...
python create_icons.py
if errorlevel 1 (
    echo ERROR: Failed to generate icons
    pause
    exit /b 1
)

REM Clean previous builds
echo.
echo Step 2: Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM Build with PyInstaller
echo.
echo Step 3: Building executable...
echo This may take a few minutes...
echo.

python -m PyInstaller batch_print.spec --clean
if errorlevel 1 (
    echo ERROR: PyInstaller build failed
    pause
    exit /b 1
)

REM Success message
echo.
echo ========================================
echo Build Complete!
echo ========================================
echo.
echo The executable has been created at:
echo   dist\HoneyBatchr\HoneyBatchr.exe
echo.
echo Optional: Register context menu? (Requires admin)
echo Run: register_context_menu.bat
echo.
pause
