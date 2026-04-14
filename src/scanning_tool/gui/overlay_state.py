from dataclasses import dataclass, field
from typing import Any, Dict

@dataclass
class OverlayState:
    # Capture overlay
    capture_overlay_root: Any = None
    capture_overlay_canvas: Any = None
    capture_rect_id: Any = None
    capture_overlay_animation_job: Any = None
    capture_overlay_last_layout: Dict[str, Any] = field(default_factory=lambda: {
        "overlay_width": None,
        "overlay_height": None,
        "left": None,
        "top": None,
        "cap_w": None,
        "cap_h": None,
    })

    # Info overlay
    info_overlay_root: Any = None
    info_overlay_canvas: Any = None
    info_text_id: Any = None
    info_overlay_geometry: Dict[str, Any] = field(default_factory=lambda: {
        "screen_width": None,
        "screen_height": None,
        "width": 0,
        "height": 0,
    })
    overlay_text: str = ""
    last_overlay_time: float = 0

    # Anchor overlay
    anchor_overlay_root: Any = None
    anchor_overlay_canvas: Any = None
    anchor_rect_id: Any = None

    border_canvas: Any = None
