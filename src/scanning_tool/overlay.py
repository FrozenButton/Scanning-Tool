"""Overlay compatibility module.

This shim keeps the old top-level overlay interface while delegating implementation
into scanning_tool.gui.overlays.
"""

from typing import Dict, Optional, TypedDict

from scanning_tool.gui.overlays import (
    choose_label_color,
    create_overlay_window,
    enforce_topmost,
    hide_anchor_overlay,
    register_anchor_sliders,
    register_capture_sliders,
    register_overlay_sliders,
    reposition_info_overlay,
    show_anchor_overlay,
    show_capture_overlay,
    show_info_overlay,
    show_overlay,
    start_capture_overlay_animation,
    start_label_timeout,
    stop_capture_overlay_animation,
    sync_anchor_sliders,
    sync_capture_sliders,
    sync_overlay_sliders,
    toggle_border,
    update_capture_overlay_region,
    update_overlay_label,
    update_anchor_overlay_region,
    update_overlay_region,
)

from scanning_tool.state import app_state


class OverlayInfo(TypedDict, total=False):
    name: str
    deposits: int


__all__ = [
    "OverlayInfo",
    "choose_label_color",
    "create_overlay_window",
    "enforce_topmost",
    "hide_anchor_overlay",
    "register_anchor_sliders",
    "register_capture_sliders",
    "register_overlay_sliders",
    "reposition_info_overlay",
    "show_anchor_overlay",
    "show_capture_overlay",
    "show_info_overlay",
    "show_overlay",
    "start_capture_overlay_animation",
    "start_label_timeout",
    "stop_capture_overlay_animation",
    "sync_anchor_sliders",
    "toggle_border",
    "sync_capture_sliders",
    "sync_overlay_sliders",
    "update_capture_overlay_region",
    "update_overlay_label",
    "update_anchor_overlay_region",
    "update_overlay_region",
    "app_state",
]
