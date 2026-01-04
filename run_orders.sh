#!/bin/bash
# Disney Magnet Order Processor - Quick Run Script
# Usage: ./run_orders.sh your_orders.csv

echo "========================================"
echo "Disney Magnet Order Processor"
echo "========================================"
echo ""

# Check if a CSV file was provided
if [ -z "$1" ]; then
    echo "No CSV file provided!"
    echo ""
    echo "USAGE:"
    echo "  ./run_orders.sh your_orders.csv"
    echo ""
    echo "Looking for CSV files in current directory..."
    echo ""
    ls -1 *.csv 2>/dev/null || echo "No CSV files found"
    echo ""
    exit 1
fi

# Check if file exists
if [ ! -f "$1" ]; then
    echo "ERROR: File not found: $1"
    echo ""
    exit 1
fi

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

echo "Processing: $1"
echo ""

# Run the script
$PYTHON_CMD process_orders.py "$1"

echo ""
echo "========================================"
echo "Processing complete!"
echo ""
echo "Check outputs/ folder for images"
echo "Check current directory for PDFs"
echo "========================================"
echo ""



