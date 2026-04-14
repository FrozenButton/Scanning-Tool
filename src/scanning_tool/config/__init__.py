"""Configuration package wrapper for the scanning tool."""

from pathlib import Path

from .loader import (
    CONFIG_FILE,
    PROJECT_ROOT,
    ROCK_TYPE_FILE,
    ROCK_TYPE_FILENAME,
    ensure_anchor_directory,
    load_config,
    resource_path,
    save_config,
)

__all__ = [
    "CONFIG_FILE",
    "PROJECT_ROOT",
    "ROCK_TYPE_FILE",
    "ROCK_TYPE_FILENAME",
    "ensure_anchor_directory",
    "load_config",
    "resource_path",
    "save_config",
]
