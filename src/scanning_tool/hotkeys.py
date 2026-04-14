"""Global hotkey listener."""

import logging

import keyboard

from scanning_tool.scanning import capture_once, toggle_continuous
from scanning_tool.gui.overlays import toggle_border

logger = logging.getLogger("scanning_tool")


def hotkey_listener() -> None:
    """Set up hotkey listeners with cross-platform error handling."""
    try:
        keyboard.add_hotkey("7", capture_once)
        keyboard.add_hotkey("ctrl+7", toggle_continuous)
        keyboard.add_hotkey("8", toggle_border)
        logger.info("Hotkeys registered: '7' for single scan, 'Ctrl+7' for continuous toggle, '8' for border toggle")
        keyboard.wait()
    except Exception as e:
        logger.warning(f"Could not set up global hotkeys: {e}")
        logger.info("Note: Linux Support is being tested.")
