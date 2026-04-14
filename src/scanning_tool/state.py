"""Centralised mutable application state, replacing module-level globals."""

from __future__ import annotations

import re
import subprocess
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np

try:
    import tkinter as tk
    from tkinter import ttk
except ImportError:  # headless environments
    tk = None  # type: ignore[assignment]
    ttk = None  # type: ignore[assignment]

ScaleWidget = Union["tk.Scale", "ttk.Scale"]


@dataclass
class AppState:
    """Single source of truth for all mutable runtime state."""

    # --- Config (persisted to config.json) ---
    cap_region: Dict[str, int] = field(
        default_factory=lambda: {"left": 1260, "top": 310, "width": 160, "height": 30}
    )
    anchor_region: Dict[str, int] = field(
        default_factory=lambda: {"left": 1100, "top": 240, "width": 320, "height": 140}
    )
    anchor_offset: Dict[str, int] = field(
        default_factory=lambda: {"x": 36, "y": 56}
    )
    anchor_threshold: float = 0.82
    auto_align_enabled: bool = True
    anchor_template_dir: str = "assets/anchor_templates"
    alignment_poll_interval_ms: int = 500
    continuous_capture_interval: float = 2.0
    info_overlay_offset: Dict[str, int] = field(
        default_factory=lambda: {"x": 0, "y": 0}
    )
    label_color: str = "yellow"
    ollama_model: str = ""
    # Fallbacks removed: always use qwen2.5vl:3b as default
    configured_ollama_host: str = ""
    default_ollama_host: str = "http://127.0.0.1:11434"
    min_confidence: float = 0.65
    debug_show_overlay: bool = True

    # --- Runtime ---
    last_result: Dict[str, Any] = field(
        default_factory=lambda: {
            "code": None,
            "code_raw": None,
            "info": None,
            "confidence": 0.0,
            "raw_text": "",
        }
    )
    last_alignment_info: Dict[str, Any] = field(
        default_factory=lambda: {
            "enabled": True,
            "matched": False,
            "template": None,
            "score": 0.0,
            "match_left": None,
            "match_top": None,
            "capture_left": None,
            "capture_top": None,
        }
    )
    continuous_mode: bool = False
    show_border: bool = True
    anchor_overlay_visible: bool = True
    anchor_tracker: Any = None  # AnchorRegionTracker instance

    # Ollama client cache
    ollama_client: Any = None
    ollama_client_host: str = ""
    ollama_server_process: Optional[subprocess.Popen] = field(default=None, repr=False)

    # Regex (compiled once)
    code_re: re.Pattern = field(
        default_factory=lambda: re.compile(
            r"(?:[A-Za-z]?-?\d[\d,\.]{1,10}|\d{2,10})", re.IGNORECASE
        )
    )
    host_scheme_re: re.Pattern = field(
        default_factory=lambda: re.compile(r"^[a-zA-Z][a-zA-Z0-9+.-]*://")
    )

    # --- Overlay windows ---
    capture_overlay_root: Any = None
    capture_overlay_canvas: Any = None
    capture_rect_id: Any = None
    capture_overlay_animation_job: Any = None
    capture_overlay_last_layout: Dict[str, Any] = field(
        default_factory=lambda: {
            "overlay_width": None,
            "overlay_height": None,
            "left": None,
            "top": None,
            "cap_w": None,
            "cap_h": None,
        }
    )

    info_overlay_root: Any = None
    info_overlay_canvas: Any = None
    info_text_id: Any = None
    info_overlay_geometry: Dict[str, Any] = field(
        default_factory=lambda: {
            "screen_width": None,
            "screen_height": None,
            "width": 0,
            "height": 0,
        }
    )
    overlay_text: str = ""
    last_overlay_time: float = 0

    anchor_overlay_root: Any = None
    anchor_overlay_canvas: Any = None
    anchor_rect_id: Any = None

    border_canvas: Any = None

    # --- GUI slider sync ---
    gui_control_state: Dict[str, Any] = field(
        default_factory=lambda: {
            "capture": {"left": None, "top": None, "width": None, "height": None},
            "anchor": {
                "left": None,
                "top": None,
                "width": None,
                "height": None,
                "offset_x": None,
                "offset_y": None,
            },
            "overlay": {"offset_x": None, "offset_y": None},
            "syncing": {"capture": False, "anchor": False, "overlay": False},
        }
    )

    # --- Deposit data (loaded at startup) ---
    rock_data: Dict = field(default_factory=dict)
    deposit_tables: Dict = field(default_factory=dict)


# Module-level singleton imported by all other modules
app_state = AppState()
