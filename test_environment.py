#!/usr/bin/env python3
"""
Test script to verify the environment setup is working correctly.
This script tests all the major dependencies without launching the full GUI.
"""

import sys
import os

def test_python_version():
    """Test Python version compatibility."""
    print(f"Python version: {sys.version}")
    version_info = sys.version_info
    
    if version_info.major < 3 or (version_info.major == 3 and version_info.minor < 8):
        print("❌ ERROR: Python 3.8+ required")
        return False
    else:
        print("✅ Python version OK")
        return True

def test_imports():
    """Test that all required packages can be imported."""
    required_packages = [
        'PIL',
        'cv2', 
        'numpy',
        'mss',
        'flask',
        'keyboard',
        'ollama',
        'tkinter'
    ]
    
    failed_imports = []
    
    for package in required_packages:
        try:
            __import__(package)
            print(f"✅ {package} imported successfully")
        except ImportError as e:
            print(f"❌ Failed to import {package}: {e}")
            failed_imports.append(package)
    
    return len(failed_imports) == 0

def test_file_structure():
    """Test that required files exist."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    required_files = [
        'scan_deposits.py',
        'requirements.txt',
        'RockTypes_2025-09-16.json',
        'config.json'  # This might not exist initially, that's OK
    ]
    
    missing_files = []
    
    for filename in required_files:
        filepath = os.path.join(script_dir, filename)
        if os.path.exists(filepath):
            print(f"✅ {filename} found")
        else:
            if filename == 'config.json':
                print(f"⚠️  {filename} not found (will be created on first run)")
            else:
                print(f"❌ {filename} missing")
                missing_files.append(filename)
    
    return len(missing_files) == 0

def test_ollama_availability():
    """Test if Ollama is available on the system."""
    import shutil
    
    if shutil.which("ollama"):
        print("✅ Ollama found in PATH")
        try:
            import subprocess
            result = subprocess.run(["ollama", "--version"], 
                                  capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                print(f"✅ Ollama version: {result.stdout.strip()}")
                return True
            else:
                print("⚠️  Ollama found but not responding correctly")
                return False
        except Exception as e:
            print(f"⚠️  Ollama found but error checking version: {e}")
            return False
    else:
        print("⚠️  Ollama not found in PATH (install from https://ollama.com/)")
        return False

def main():
    """Run all tests."""
    print("=== Environment Test Script ===")
    print()
    
    tests = [
        ("Python Version", test_python_version),
        ("Package Imports", test_imports),
        ("File Structure", test_file_structure),
        ("Ollama Availability", test_ollama_availability)
    ]
    
    results = []
    
    for test_name, test_func in tests:
        print(f"--- Testing {test_name} ---")
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"❌ Test failed with exception: {e}")
            results.append((test_name, False))
        print()
    
    # Summary
    print("=== Test Summary ===")
    all_passed = True
    for test_name, passed in results:
        status = "✅ PASS" if passed else "❌ FAIL"
        print(f"{status}: {test_name}")
        if not passed:
            all_passed = False
    
    print()
    if all_passed:
        print("🎉 All tests passed! The environment is ready.")
        return 0
    else:
        print("⚠️  Some tests failed. Please check the issues above.")
        return 1

if __name__ == "__main__":
    sys.exit(main())