#!/bin/bash

# Star Citizen Scanning Tool - Linux Launch Script
# This script sets up a virtual environment and launches the application

set -e  # Exit on any error

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$SCRIPT_DIR/venv"
PYTHON_CMD="python3"

echo "=== Star Citizen Scanning Tool - Linux Setup ==="
echo "Script directory: $SCRIPT_DIR"

# Check if Python 3 is installed
if ! command -v $PYTHON_CMD &> /dev/null; then
    echo "Error: Python 3 is not installed or not in PATH"
    echo "Please install Python 3.8 or newer:"
    echo "  Ubuntu/Debian: sudo apt update && sudo apt install python3 python3-pip python3-venv"
    echo "  Fedora/RHEL: sudo dnf install python3 python3-pip"
    echo "  Arch: sudo pacman -S python python-pip"
    exit 1
fi

# Check Python version
PYTHON_VERSION=$($PYTHON_CMD --version 2>&1 | cut -d' ' -f2)
PYTHON_MAJOR=$(echo $PYTHON_VERSION | cut -d'.' -f1)
PYTHON_MINOR=$(echo $PYTHON_VERSION | cut -d'.' -f2)

echo "Found Python $PYTHON_VERSION"

if [ "$PYTHON_MAJOR" -lt 3 ] || ([ "$PYTHON_MAJOR" -eq 3 ] && [ "$PYTHON_MINOR" -lt 8 ]); then
    echo "Error: Python 3.8 or newer is required (found $PYTHON_VERSION)"
    exit 1
fi

# Check if virtual environment exists
if [ ! -d "$VENV_DIR" ]; then
    echo "Creating virtual environment..."
    $PYTHON_CMD -m venv "$VENV_DIR"
    echo "Virtual environment created at: $VENV_DIR"
else
    echo "Virtual environment already exists at: $VENV_DIR"
fi

# Activate virtual environment
echo "Activating virtual environment..."
source "$VENV_DIR/bin/activate"

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip

# Install requirements if they don't exist or if requirements.txt is newer
if [ ! -f "$VENV_DIR/.requirements_installed" ] || [ "$SCRIPT_DIR/requirements.txt" -nt "$VENV_DIR/.requirements_installed" ]; then
    echo "Installing Python requirements..."
    pip install -r "$SCRIPT_DIR/requirements.txt"
    touch "$VENV_DIR/.requirements_installed"
    echo "Requirements installed successfully"
else
    echo "Requirements already up to date"
fi

# Check if Ollama is installed
if ! command -v ollama &> /dev/null; then
    echo ""
    echo "WARNING: Ollama is not installed!"
    echo "The application will prompt you to install it from https://ollama.com/"
    echo "For Linux, you can install it with:"
    echo "  curl -fsSL https://ollama.com/install.sh | sh"
    echo ""
fi

# Launch the application
echo "Launching Star Citizen Scanning Tool..."
echo "Press Ctrl+C to stop the application"
echo ""

cd "$SCRIPT_DIR"
python scan_deposits.py

echo ""
echo "Application closed."