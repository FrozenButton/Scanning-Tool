"""Common overlay utilities and Tk helper functions."""

from loguru import logger
from typing import Callable, Optional, TypeVar

import tkinter as tk

T = TypeVar("T")

TOPMOST_INTERVAL_MS = 1500
CAPTURE_ANIMATION_INTERVAL_MS = 33
INFO_LABEL_TIMEOUT_MS = 500
INFO_LABEL_CLEAR_AFTER_S = 10
CAPTURE_OVERLAY_PADDING_X = 100
CAPTURE_OVERLAY_PADDING_Y = 40
ANCHOR_OVERLAY_PAD = 40
MIN_INFO_OVERLAY_WIDTH = 400
MAX_INFO_OVERLAY_WIDTH = 800
INFO_OVERLAY_HEIGHT = 120
SCREEN_MARGIN = 40


def safe_tk(func: Callable[[], T], default: Optional[T] = None) -> Optional[T]:
    try:
        return func()
    except tk.TclError:
        return default


def enforce_topmost(window: tk.Toplevel, interval_ms: int = TOPMOST_INTERVAL_MS) -> None:
    if safe_tk(window.winfo_exists, False) is False:
        return
    try:
        window.attributes("-topmost", True)  # type: ignore[attr-defined]
        window.lift()  # type: ignore[attr-defined]
    except tk.TclError:
        return
    window.after(interval_ms, lambda: enforce_topmost(window, interval_ms))


def create_overlay_window(width: int, height: int, left: int, top: int) -> tk.Toplevel:
    window = tk.Toplevel()
    window.attributes("-transparentcolor", "black")  # type: ignore[attr-defined]
    window.attributes("-topmost", True)  # type: ignore[attr-defined]
    window.overrideredirect(True)
    window.configure(bg="black")
    window.geometry(f"{width}x{height}+{left}+{top}")
    enforce_topmost(window)
    return window
