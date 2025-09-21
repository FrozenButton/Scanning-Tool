# Launch Scripts Summary

This document provides details about the launch scripts created for the Star Citizen Scanning Tool.

## Created Files

### Launch Scripts
1. **`launch_linux.sh`** - Linux/Unix launch script
2. **`launch_windows.bat`** - Windows batch script  
3. **`launch_windows.ps1`** - Windows PowerShell script (alternative)
4. **`launch_macos.sh`** - macOS launch script

### Support Files
5. **`test_environment.py`** - Environment testing script
6. **`.gitignore`** - Git ignore file for virtual environments and temporary files

## Features

### Windows Scripts (`launch_windows.bat` & `launch_windows.ps1`)

**Key Features:**
- 🐍 **Portable Python Installation**: Automatically downloads and installs Python 3.13 embedded distribution
- 📦 **No System Python Required**: Creates a completely isolated, portable environment
- 🔧 **Automatic Setup**: Creates virtual environment and installs dependencies
- 💾 **Smart Caching**: Only re-downloads/reinstalls when needed
- 🛡️ **Error Handling**: Comprehensive error checking and user guidance

**How it Works:**
1. Downloads Python 3.13 embedded distribution (~9MB) to `./python/` folder
2. Configures embedded Python to support pip
3. Creates virtual environment in `./venv/` folder
4. Installs requirements from `requirements.txt`
5. Launches the application

**Usage:**
```cmd
# Batch version (recommended)
launch_windows.bat

# PowerShell version (alternative)
powershell -ExecutionPolicy Bypass -File launch_windows.ps1
```

### Linux Script (`launch_linux.sh`)

**Key Features:**
- 🔍 **Python Detection**: Automatically finds system Python 3.8+
- 📋 **Dependency Guidance**: Provides install commands for different distributions
- 🔧 **Virtual Environment**: Creates isolated environment
- ⚡ **Fast Setup**: Reuses existing installations when possible

**Requirements:**
- Python 3.8+ installed on system
- pip and venv modules available

**Usage:**
```bash
./launch_linux.sh
```

## Environment Testing

The `test_environment.py` script can be used to verify the environment setup:

```bash
python test_environment.py
```

**Tests Performed:**
- ✅ Python version compatibility (3.8+)
- ✅ Required package imports
- ✅ File structure verification
- ✅ Ollama availability check

## Directory Structure After Setup

```
Scanning-Tool/
├── launch_linux.sh          # Linux launcher
├── launch_windows.bat       # Windows launcher (batch)
├── launch_windows.ps1       # Windows launcher (PowerShell)
├── launch_macos.sh          # macOS launcher
├── test_environment.py      # Environment tester
├── scan_deposits.py         # Main application
├── requirements.txt         # Python dependencies
├── config.json             # App config (created on first run)
├── RockTypes_2025-09-16.json # Rock data
├── templates/              # Flask templates
├── venv/                   # Virtual environment (auto-created)
└── python/                 # Portable Python (Windows only, auto-created)
```

## User Instructions

### For End Users

**Windows Users:**
1. Download the project
2. Double-click `launch_windows.bat`
3. Wait for automatic setup to complete
4. Application launches automatically

**Linux/macOS Users:**
1. Download the project
2. Open terminal in project directory
3. Run `./launch_linux.sh` or `./launch_macos.sh`
4. Follow any Python installation prompts if needed
5. Application launches automatically

### For Developers

**Testing Environment:**
```bash
python test_environment.py
```

**Manual Virtual Environment:**
```bash
python3 -m venv venv
source venv/bin/activate  # Linux/macOS
# OR
venv\Scripts\activate     # Windows
pip install -r requirements.txt
```

## Technical Details

### Windows Portable Python
- Downloads from: `https://www.python.org/ftp/python/3.13.0/python-3.13.0-embed-amd64.zip`
- Size: ~9MB download
- No registry modifications
- Completely portable and isolated
- Includes pip support after configuration

### Virtual Environment Management
- Creates isolated Python environments
- Prevents conflicts with system packages
- Automatic requirement caching with timestamp checking
- Cross-platform activation handling

### Error Handling
- Network connectivity checks
- Permission validation
- Clear error messages with solutions
- Graceful fallbacks where possible

## Troubleshooting

### Windows
- **PowerShell Execution Policy**: Use `powershell -ExecutionPolicy Bypass -File launch_windows.ps1`
- **Antivirus Blocking**: Whitelist the project folder
- **Network Issues**: Check firewall settings for Python downloads

### Linux
- **Python Not Found**: Install with package manager (apt, dnf, pacman, etc.)
- **Permission Denied**: Run `chmod +x launch_linux.sh`
- **Virtual Environment Issues**: Install `python3-venv` package

### macOS
- **Python Not Found**: Install from python.org or use Homebrew
- **Permission Denied**: Run `chmod +x launch_macos.sh`
- **Gatekeeper Issues**: Right-click → Open for first run