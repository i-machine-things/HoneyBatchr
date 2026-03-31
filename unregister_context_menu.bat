@echo off
REM Context Menu Unregistration Script for Honey Batchr
REM This script removes "Batch Print with Honey Batchr" from the Windows context menu
REM Requires Administrator privileges

echo.
echo ========================================
echo Honey Batchr - Context Menu Removal
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

echo Removing context menu entries...

REM Remove from file context menu
reg delete "HKEY_CLASSES_ROOT\*\shell\Batch Print with Honey Batchr" /f >nul 2>&1

REM Remove from folder context menu
reg delete "HKEY_CLASSES_ROOT\Folder\shell\Batch Print with Honey Batchr" /f >nul 2>&1

REM Show success message
echo.
echo ========================================
echo Removal Complete!
echo ========================================
echo.
echo The context menu entry has been removed.
echo.
echo To re-add it, run: register_context_menu.bat
echo.
pause
