#!/bin/bash
# Disney Magnet Order Processor - GUI Launcher
# Run this file to start the application

echo "========================================"
echo "Disney Magnet Order Processor"
echo "Starting GUI Application..."
echo "========================================"
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null && ! command -v python &> /dev/null; then
    echo "ERROR: Python is not installed"
    echo "Please install Python 3.7 or higher"
    echo ""
    exit 1
fi

# Use python3 if available, otherwise python
PYTHON_CMD="python3"
if ! command -v python3 &> /dev/null; then
    PYTHON_CMD="python"
fi

# Check if required packages are installed
$PYTHON_CMD -c "import PIL, PyPDF2" &> /dev/null
if [ $? -ne 0 ]; then
    echo "Installing required packages..."
    $PYTHON_CMD -m pip install -r requirements.txt
    echo ""
fi

# Launch GUI
echo "Starting application..."
$PYTHON_CMD gui_app.py

if [ $? -ne 0 ]; then
    echo ""
    echo "========================================"
    echo "Application closed with errors"
    echo "========================================"
    read -p "Press Enter to continue..."
fi

