#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "=== Star Citizen Scanning Tool - Git Bash Windows Launch ==="
echo "Project directory: $SCRIPT_DIR"

if ! command -v uv >/dev/null 2>&1; then
  echo "Error: uv is not installed or not available in PATH."
  echo "Install uv or use Windows CMD/PowerShell with launch_windows.bat."
  exit 1
fi

cd "$SCRIPT_DIR"
echo "Using uv-managed Python to launch the app."
uv run --managed-python --with-requirements requirements.txt -- python scan_deposits.py
