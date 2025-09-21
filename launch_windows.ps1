# Star Citizen Scanning Tool - Windows PowerShell Launch Script
# This script sets up a portable Python environment and launches the application

param(
    [switch]$SkipPython,
    [switch]$Force
)

Write-Host "=== Star Citizen Scanning Tool - Windows Setup ===" -ForegroundColor Cyan

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonDir = Join-Path $ScriptDir "python"
$VenvDir = Join-Path $ScriptDir "venv"
$PythonExe = Join-Path $PythonDir "python.exe"
$PythonVersion = "3.13.7"
$PythonUrl = "https://www.python.org/ftp/python/3.13.7/python-3.13.7-embed-amd64.zip"

Write-Host "Script directory: $ScriptDir"

# Function to download with progress
function Get-File {
    param($Url, $Output)
    try {
        $WebClient = New-Object System.Net.WebClient
        $WebClient.DownloadFile($Url, $Output)
        return $true
    }
    catch {
        Write-Host "Download failed: $_" -ForegroundColor Red
        return $false
    }
}

# Check if portable Python exists or if forced reinstall
if ((-not (Test-Path $PythonExe)) -or $Force) {
    if (-not $SkipPython) {
        Write-Host "Installing portable Python $PythonVersion..." -ForegroundColor Yellow
        
        # Create python directory
        if (-not (Test-Path $PythonDir)) {
            New-Item -ItemType Directory -Path $PythonDir | Out-Null
        }
        
        $ZipPath = Join-Path $ScriptDir "python.zip"
        
        # Download Python embedded distribution
        Write-Host "Downloading Python $PythonVersion embedded distribution..."
        Write-Host "This may take a few minutes depending on your internet connection..."
        
        if (Get-File $PythonUrl $ZipPath) {
            Write-Host "Download completed!" -ForegroundColor Green
            
            # Extract Python
            Write-Host "Extracting Python..."
            Expand-Archive -Path $ZipPath -DestinationPath $PythonDir -Force
            
            # Clean up zip file
            Remove-Item $ZipPath
            
            # Enable pip by modifying python313._pth
            $PthFile = Join-Path $PythonDir "python313._pth"
            if (Test-Path $PthFile) {
                Write-Host "Configuring Python for pip support..."
                $PthLines = Get-Content $PthFile
                $PthLines = $PthLines | ForEach-Object { $_ -replace '^#import site', 'import site' }
                Set-Content $PthFile -Value $PthLines
            }
            
            # Install pip
            Write-Host "Installing pip..."
            & $PythonExe -m ensurepip --default-pip
            
            Write-Host "Portable Python installation completed!" -ForegroundColor Green
        }
        else {
            Write-Host "Failed to download Python. Please check your internet connection." -ForegroundColor Red
            Write-Host "You can manually download from: $PythonUrl"
            Write-Host "Extract to: $PythonDir"
            Read-Host "Press Enter to exit"
            exit 1
        }
    }
}
else {
    Write-Host "Portable Python found at: $PythonExe" -ForegroundColor Green
}

# Check Python version
Write-Host "Checking Python version..."
$CurrentVersion = (& $PythonExe --version 2>&1) -replace "Python ", ""
Write-Host "Found Python $CurrentVersion" -ForegroundColor Green

# Check if virtual environment exists
if ((-not (Test-Path $VenvDir)) -or $Force) {
    Write-Host "Creating virtual environment..." -ForegroundColor Yellow
    & $PythonExe -m venv $VenvDir
    Write-Host "Virtual environment created at: $VenvDir" -ForegroundColor Green
}
else {
    Write-Host "Virtual environment already exists at: $VenvDir" -ForegroundColor Green
}

# Activate virtual environment
Write-Host "Activating virtual environment..."
$ActivateScript = Join-Path $VenvDir "Scripts\Activate.ps1"
& $ActivateScript

# Upgrade pip
Write-Host "Upgrading pip..."
& "$VenvDir\Scripts\python.exe" -m pip install --upgrade pip

# Install requirements
$ReqMarker = Join-Path $VenvDir ".requirements_installed"
$RequirementsTxt = Join-Path $ScriptDir "requirements.txt"

$InstallRequired = $false
if (-not (Test-Path $ReqMarker)) {
    $InstallRequired = $true
}
elseif ((Get-Item $RequirementsTxt).LastWriteTime -gt (Get-Item $ReqMarker).LastWriteTime) {
    $InstallRequired = $true
}

if ($InstallRequired -or $Force) {
    Write-Host "Installing Python requirements..." -ForegroundColor Yellow
    & "$VenvDir\Scripts\pip.exe" install -r $RequirementsTxt
    New-Item -ItemType File -Path $ReqMarker -Force | Out-Null
    Write-Host "Requirements installed successfully" -ForegroundColor Green
}
else {
    Write-Host "Requirements already up to date" -ForegroundColor Green
}

# Check if Ollama is installed
if (-not (Get-Command ollama -ErrorAction SilentlyContinue)) {
    Write-Host ""
    Write-Host "WARNING: Ollama is not installed!" -ForegroundColor Red
    Write-Host "The application will prompt you to install it from https://ollama.com/"
    Write-Host "For Windows, download the installer from the official website."
    Write-Host ""
}

# Launch the application
Write-Host "Launching Star Citizen Scanning Tool..." -ForegroundColor Cyan
Write-Host "Press Ctrl+C to stop the application"
Write-Host ""

Set-Location $ScriptDir
& "$VenvDir\Scripts\python.exe" scan_deposits.py

Write-Host ""
Write-Host "Application closed." -ForegroundColor Green