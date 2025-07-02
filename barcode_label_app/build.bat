@echo off
echo ====================================
echo Barcode Label Generator - Build Script
echo ====================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or later from https://python.org
    pause
    exit /b 1
)

echo Python found. Starting build process...
echo.

REM Run the build script
python build_exe.py

echo.
echo Build process completed!
echo Check the output above for any errors.
echo.
pause 