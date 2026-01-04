@echo off
REM Disney Magnet Order Processor - Quick Run Script
REM Usage: Drag and drop your CSV file onto this batch file

echo ========================================
echo Disney Magnet Order Processor
echo ========================================
echo.

REM Check if a file was dragged onto this batch file
if "%~1"=="" (
    echo No CSV file provided!
    echo.
    echo USAGE:
    echo   1. Drag and drop your CSV file onto this batch file
    echo   OR
    echo   2. Run from command line: run_orders.bat your_orders.csv
    echo.
    echo Looking for CSV files in current directory...
    echo.
    dir /b *.csv
    echo.
    pause
    exit /b 1
)

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo Processing: %~1
echo.

REM Run the script
python process_orders.py "%~1"

echo.
echo ========================================
echo Processing complete!
echo.
echo Check outputs\ folder for images
echo Check current directory for PDFs
echo ========================================
echo.
pause



