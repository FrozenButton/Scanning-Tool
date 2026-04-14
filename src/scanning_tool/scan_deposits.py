"""Backward-compatible entry point for launch scripts."""
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))
from scanning_tool.main import main

if __name__ == "__main__":
    main()