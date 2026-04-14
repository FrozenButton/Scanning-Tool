"""Pure overlay geometry helpers."""

from typing import Dict, Tuple

from .base import (
    ANCHOR_OVERLAY_PAD,
    CAPTURE_OVERLAY_PADDING_X,
    CAPTURE_OVERLAY_PADDING_Y,
    INFO_OVERLAY_HEIGHT,
    MAX_INFO_OVERLAY_WIDTH,
    MIN_INFO_OVERLAY_WIDTH,
    SCREEN_MARGIN,
)
from scanning_tool.state import app_state


def compute_info_overlay_geometry(screen_width: int, screen_height: int) -> Tuple[int, int, int, int]:
    overlay_width = max(
        MIN_INFO_OVERLAY_WIDTH,
        min(MAX_INFO_OVERLAY_WIDTH, screen_width - SCREEN_MARGIN),
    )
    overlay_height = INFO_OVERLAY_HEIGHT
    base_left = max(0, (screen_width - overlay_width) // 2)
    base_top = max(0, int(screen_height * 0.35) - overlay_height // 2)

    offset_x = int(app_state.info_overlay_offset.get("x", 0))
    offset_y = int(app_state.info_overlay_offset.get("y", 0))

    max_left = max(0, screen_width - overlay_width)
    max_top = max(0, screen_height - overlay_height)

    left = min(max(0, base_left + offset_x), max_left)
    top = min(max(0, base_top + offset_y), max_top)
    return overlay_width, overlay_height, left, top


def compute_capture_overlay_layout() -> Dict[str, int]:
    cap_w = int(app_state.cap_region["width"])
    cap_h = int(app_state.cap_region["height"])

    overlay_width = cap_w + CAPTURE_OVERLAY_PADDING_X
    overlay_height = cap_h + CAPTURE_OVERLAY_PADDING_Y
    left = int(app_state.cap_region["left"]) - (CAPTURE_OVERLAY_PADDING_X // 2)
    top = int(app_state.cap_region["top"]) - CAPTURE_OVERLAY_PADDING_Y

    return {
        "overlay_width": overlay_width,
        "overlay_height": overlay_height,
        "left": left,
        "top": top,
        "padding_x": CAPTURE_OVERLAY_PADDING_X,
        "padding_y": CAPTURE_OVERLAY_PADDING_Y,
        "cap_w": cap_w,
        "cap_h": cap_h,
    }


def compute_anchor_overlay_geometry() -> Dict[str, int]:
    width = int(app_state.anchor_region["width"]) + ANCHOR_OVERLAY_PAD
    height = int(app_state.anchor_region["height"]) + ANCHOR_OVERLAY_PAD
    left = int(app_state.anchor_region["left"]) - (ANCHOR_OVERLAY_PAD // 2)
    top = int(app_state.anchor_region["top"]) - (ANCHOR_OVERLAY_PAD // 2)

    return {
        "width": width,
        "height": height,
        "left": left,
        "top": top,
    }
