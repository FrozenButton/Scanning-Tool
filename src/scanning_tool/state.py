"""Centralised application state for the scanning tool."""

from __future__ import annotations

from typing import Union

try:
    import tkinter as tk
    from tkinter import ttk
except ImportError:  # headless environments
    tk = None  # type: ignore[assignment]
    ttk = None  # type: ignore[assignment]

from .context import AppContext, app_state as _context_app_state

ScaleWidget = Union[tk.Scale, ttk.Scale]

app_state = _context_app_state

__all__ = ["AppContext", "app_state", "ScaleWidget"]
