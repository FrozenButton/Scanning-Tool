from dataclasses import dataclass, field
from typing import Any, Dict, Optional

import tkinter as tk

@dataclass
class OverlayState:
    # Capture overlay
    capture_overlay_root: Optional[tk.Toplevel] = None
    capture_overlay_canvas: Optional[tk.Canvas] = None
    capture_rect_id: Optional[int] = None
    capture_overlay_animation_job: Optional[str] = None
    capture_overlay_last_layout: Dict[str, Any] = field(default_factory=lambda: {
        "overlay_width": None,
        "overlay_height": None,
        "left": None,
        "top": None,
        "cap_w": None,
        "cap_h": None,
    })

    # Info overlay
    info_overlay_root: Optional[tk.Toplevel] = None
    info_overlay_canvas: Optional[tk.Canvas] = None
    info_text_id: Optional[int] = None
    info_overlay_geometry: Dict[str, Any] = field(default_factory=lambda: {
        "screen_width": None,
        "screen_height": None,
        "width": 0,
        "height": 0,
    })
    overlay_text: str = ""
    last_overlay_time: float = 0

    # Anchor overlay
    anchor_overlay_root: Optional[tk.Toplevel] = None
    anchor_overlay_canvas: Optional[tk.Canvas] = None
    anchor_rect_id: Optional[int] = None
    anchor_overlay_visible: bool = True

    border_canvas: Optional[tk.Canvas] = None
    show_border: bool = True
