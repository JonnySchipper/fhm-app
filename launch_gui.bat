@echo off
REM Disney Magnet Order Processor - GUI Launcher
REM Double-click this file to start the application

echo ========================================
echo Disney Magnet Order Processor
echo Starting GUI Application...
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

REM Check if required packages are installed
python -c "import PIL, PyPDF2" >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    pip install -r requirements.txt
    echo.
)

REM Launch GUI
echo Starting application...
python gui_app.py

if errorlevel 1 (
    echo.
    echo ========================================
    echo Application closed with errors
    echo ========================================
    pause
)

