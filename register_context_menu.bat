@echo off
REM Context Menu Registration Script for Honey Batchr
REM This script adds "Batch Print with Honey Batchr" to the Windows context menu
REM Requires Administrator privileges

echo.
echo ========================================
echo Honey Batchr - Context Menu Registration
echo ========================================
echo.

REM Check for admin privileges
net session >nul 2>&1
if errorlevel 1 (
    echo ERROR: This script requires Administrator privileges
    echo.
    echo Please right-click this file and select "Run as administrator"
    pause
    exit /b 1
)

REM Get the path to HoneyBatchr.exe
setlocal enabledelayedexpansion

REM Try to find in dist folder
if exist "dist\HoneyBatchr\HoneyBatchr.exe" (
    set "APP_PATH=!CD!\dist\HoneyBatchr\HoneyBatchr.exe"
) else if exist "HoneyBatchr.exe" (
    set "APP_PATH=!CD!\HoneyBatchr.exe"
) else (
    echo ERROR: HoneyBatchr.exe not found
    echo.
    echo Please build the application first by running: build.bat
    pause
    exit /b 1
)

echo Found application at:
echo !APP_PATH!
echo.

REM Create registry entries
echo Adding context menu entry...

REM Add to file context menu
reg add "HKEY_CLASSES_ROOT\*\shell\Batch Print with Honey Batchr" /v "" /d "Batch Print with Honey Batchr" /f >nul
reg add "HKEY_CLASSES_ROOT\*\shell\Batch Print with Honey Batchr\command" /v "" /d "\"!APP_PATH!\" \"%%%%1\"" /f >nul

REM Add icon
reg add "HKEY_CLASSES_ROOT\*\shell\Batch Print with Honey Batchr" /v "Icon" /d "!APP_PATH!" /f >nul

REM Add to folder context menu
reg add "HKEY_CLASSES_ROOT\Folder\shell\Batch Print with Honey Batchr" /v "" /d "Batch Print with Honey Batchr" /f >nul
reg add "HKEY_CLASSES_ROOT\Folder\shell\Batch Print with Honey Batchr\command" /v "" /d "\"!APP_PATH!\" \"%%%%1\"" /f >nul
reg add "HKEY_CLASSES_ROOT\Folder\shell\Batch Print with Honey Batchr" /v "Icon" /d "!APP_PATH!" /f >nul

REM Show success message
echo.
echo ========================================
echo Registration Complete!
echo ========================================
echo.
echo The context menu entry has been added:
echo   "Batch Print with Honey Batchr"
echo.
echo You can now right-click any file or folder and select
echo "Batch Print with Honey Batchr" to open it in the application.
echo.
echo To remove this entry, run: unregister_context_menu.bat
echo.
pause
