@echo off
setlocal enabledelayedexpansion

REM Star Citizen Scanning Tool - Windows Launch Script
REM This script sets up a portable Python environment and launches the application

echo === Star Citizen Scanning Tool - Windows Setup ===

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "PYTHON_DIR=%SCRIPT_DIR%\python"
set "VENV_DIR=%SCRIPT_DIR%\venv"
set "PYTHON_EXE=%PYTHON_DIR%\python.exe"
set "PYTHON_VERSION=3.13.7"
set "PYTHON_URL=https://www.python.org/ftp/python/3.13.7/python-3.13.7-embed-amd64.zip"

echo Script directory: %SCRIPT_DIR%

REM Check if portable Python exists
if not exist "%PYTHON_EXE%" (
    echo Portable Python not found. Installing Python %PYTHON_VERSION%...
    
    REM Create python directory
    if not exist "%PYTHON_DIR%" mkdir "%PYTHON_DIR%"
    
    REM Download Python embedded distribution
    echo Downloading Python %PYTHON_VERSION% embedded distribution...
    echo This may take a few minutes depending on your internet connection...
    
    REM Use PowerShell to download (available on Windows 7+ by default)
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%SCRIPT_DIR%\python.zip'}"
    
    if not exist "%SCRIPT_DIR%\python.zip" (
        echo Error: Failed to download Python. Please check your internet connection.
        echo You can manually download Python from: %PYTHON_URL%
        echo Extract it to: %PYTHON_DIR%
        pause
        exit /b 1
    )
    
    REM Extract Python
    echo Extracting Python...
    powershell -Command "Expand-Archive -Path '%SCRIPT_DIR%\python.zip' -DestinationPath '%PYTHON_DIR%' -Force"
    
    REM Clean up zip file
    del "%SCRIPT_DIR%\python.zip"
    
    REM Enable pip by modifying python313._pth
    echo Configuring Python for pip support...
    if exist "%PYTHON_DIR%\python313._pth" (
        REM Remove the comment from import site line
        powershell -Command "(Get-Content '%PYTHON_DIR%\python313._pth') -replace '^#import site', 'import site' | Set-Content '%PYTHON_DIR%\python313._pth'"
    )
    
    REM Install pip manually for embedded Python
    echo Installing pip...
    "%PYTHON_EXE%" -m ensurepip --default-pip
    
    echo Portable Python installation completed!
) else (
    echo Portable Python found at: %PYTHON_EXE%
)

REM Check Python version
echo Checking Python version...
for /f "tokens=2" %%i in ('"%PYTHON_EXE%" --version 2^>^&1') do set "CURRENT_VERSION=%%i"
echo Found Python %CURRENT_VERSION%

REM Check if virtual environment exists
if not exist "%VENV_DIR%" (
    echo Creating virtual environment...
    "%PYTHON_EXE%" -m venv "%VENV_DIR%"
    echo Virtual environment created at: %VENV_DIR%
) else (
    echo Virtual environment already exists at: %VENV_DIR%
)

REM Activate virtual environment
echo Activating virtual environment...
call "%VENV_DIR%\Scripts\activate.bat"

REM Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip

REM Install requirements if they don't exist or if requirements.txt is newer
set "REQ_MARKER=%VENV_DIR%\.requirements_installed"
set "INSTALL_REQUIRED=0"

if not exist "%REQ_MARKER%" set "INSTALL_REQUIRED=1"

REM Check if requirements.txt is newer than marker file
if exist "%REQ_MARKER%" (
    for %%i in ("%SCRIPT_DIR%\requirements.txt") do set "REQ_DATE=%%~ti"
    for %%i in ("%REQ_MARKER%") do set "MARKER_DATE=%%~ti"
    REM Simple date comparison - if requirements.txt is newer, reinstall
    if "!REQ_DATE!" GTR "!MARKER_DATE!" set "INSTALL_REQUIRED=1"
)

if !INSTALL_REQUIRED! == 1 (
    echo Installing Python requirements...
    pip install -r "%SCRIPT_DIR%\requirements.txt"
    echo. > "%REQ_MARKER%"
    echo Requirements installed successfully
) else (
    echo Requirements already up to date
)

REM Check if Ollama is installed
where ollama >nul 2>nul
if errorlevel 1 (
    echo.
    echo WARNING: Ollama is not installed!
    echo The application will prompt you to install it from https://ollama.com/
    echo For Windows, download the installer from the official website.
    echo.
)

REM Launch the application
echo Launching Star Citizen Scanning Tool...
echo Press Ctrl+C to stop the application
echo.

cd /d "%SCRIPT_DIR%"
python scan_deposits.py

echo.
echo Application closed.
pause