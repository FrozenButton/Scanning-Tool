@echo off
setlocal enabledelayedexpansion

REM Build a standalone Windows executable for the Star Citizen Scanning Tool
REM using the embeddable Python distribution so end users don't need Python
REM installed. Run this script from the repository root on a Windows machine.

set "SCRIPT_DIR=%~dp0"
set "PY_VERSION=3.13.7"
set "PY_EMBED_URL=https://www.python.org/ftp/python/%PY_VERSION%/python-%PY_VERSION%-embed-amd64.zip"
set "PY_DIR=%SCRIPT_DIR%python"
set "PY_EXE=%PY_DIR%python.exe"
set "PY_ZIP=%SCRIPT_DIR%python.zip"
set "GET_PIP_URL=https://bootstrap.pypa.io/get-pip.py"
set "GET_PIP_FILE=%PY_DIR%get-pip.py"
set "PTH_FILE=%PY_DIR%python313._pth"
set "DIST_DIR=%SCRIPT_DIR%dist"
set "BUILD_DIR=%SCRIPT_DIR%build"
set "EXE_NAME=StarCitizenScanningTool"

if not exist "%PY_EXE%" (
    echo [1/6] Downloading portable Python %PY_VERSION%...
    powershell -NoProfile -Command "Invoke-WebRequest -Uri '%PY_EMBED_URL%' -OutFile '%PY_ZIP%'"
    if not exist "%PY_ZIP%" (
        echo Failed to download Python from %PY_EMBED_URL%
        exit /b 1
    )
    echo [2/6] Unpacking portable Python...
    powershell -NoProfile -Command "Expand-Archive -Path '%PY_ZIP%' -DestinationPath '%PY_DIR%' -Force"
    del "%PY_ZIP%" >nul 2>&1
)

if exist "%PTH_FILE%" (
    findstr /C:"import site" "%PTH_FILE%" >nul 2>&1
    if errorlevel 1 (
        echo import site>>"%PTH_FILE%"
    )
)

if not exist "%GET_PIP_FILE%" (
    echo [3/6] Downloading get-pip bootstrapper...
    powershell -NoProfile -Command "Invoke-WebRequest -Uri '%GET_PIP_URL%' -OutFile '%GET_PIP_FILE%'"
)

if not exist "%GET_PIP_FILE%" (
    echo Failed to download get-pip.py from %GET_PIP_URL%
    exit /b 1
)

echo [4/6] Installing pip into the portable Python...
"%PY_EXE%" "%GET_PIP_FILE%" --no-warn-script-location --disable-pip-version-check
if errorlevel 1 (
    echo Failed to bootstrap pip.
    exit /b 1
)

echo [5/6] Installing build dependencies (pyinstaller + app requirements)...
"%PY_EXE%" -m pip install --upgrade pip
"%PY_EXE%" -m pip install --upgrade pyinstaller
"%PY_EXE%" -m pip install -r "%SCRIPT_DIR%requirements.txt"
if errorlevel 1 (
    echo Dependency installation failed.
    exit /b 1
)

if exist "%DIST_DIR%" (
    rmdir /S /Q "%DIST_DIR%"
)
if exist "%BUILD_DIR%" (
    rmdir /S /Q "%BUILD_DIR%"
)

echo [6/6] Building the standalone executable...
"%PY_EXE%" -m PyInstaller --noconfirm --clean --name "%EXE_NAME%" --windowed ^
  --add-data "RockTypes_2025-09-16.json;." ^
  --add-data "templates;templates" ^
  --add-data "assets;assets" ^
  "%SCRIPT_DIR%scan_deposits.py"
if errorlevel 1 (
    echo PyInstaller build failed.
    exit /b 1
)

echo.
echo Build complete! The packaged app is at:
echo     %DIST_DIR%\%EXE_NAME%\%EXE_NAME%.exe
echo.
echo You can zip the entire dist\%EXE_NAME% folder and share it with users.

echo Done.
endlocal
