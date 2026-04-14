"""Overlay package API."""

from typing import Optional

import tkinter as tk

from .anchor import hide_anchor_overlay, show_anchor_overlay, update_anchor_overlay_region
from .base import create_overlay_window, enforce_topmost, safe_tk
from .capture import (
    hide_capture_overlay,
    show_capture_overlay,
    start_capture_overlay_animation,
    stop_capture_overlay_animation,
    update_capture_overlay_region,
)
from .geometry import compute_info_overlay_geometry
from .info import (
    choose_label_color,
    reposition_info_overlay,
    show_info_overlay,
    start_label_timeout,
    toggle_border,
    update_overlay_label,
    hide_info_overlay,
)
from .slider_sync import (
    register_anchor_sliders,
    register_capture_sliders,
    register_overlay_sliders,
    sync_anchor_sliders,
    sync_capture_sliders,
    sync_overlay_sliders,
)
from scanning_tool.state import app_state

__all__ = [
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
    "hide_capture_overlay",
    "show_info_overlay",
    "hide_info_overlay",
    "show_overlay",
    "start_capture_overlay_animation",
    "toggle_border",
    "start_label_timeout",
    "stop_capture_overlay_animation",
    "sync_anchor_sliders",
    "sync_capture_sliders",
    "sync_overlay_sliders",
    "update_capture_overlay_region",
    "update_overlay_label",
    "update_anchor_overlay_region",
    "update_overlay_region",
    "destroy_all_overlays",
]


def show_overlay(screen_width: Optional[int] = None, screen_height: Optional[int] = None) -> None:
    show_anchor_overlay()
    show_capture_overlay()

    if screen_width is None or screen_height is None:
        overlay_state = app_state.overlay_state
        if overlay_state.capture_overlay_root and safe_tk(overlay_state.capture_overlay_root.winfo_exists, False):
            screen_width = safe_tk(overlay_state.capture_overlay_root.winfo_screenwidth, 1920) or 1920
            screen_height = safe_tk(overlay_state.capture_overlay_root.winfo_screenheight, 1080) or 1080

    if screen_width is not None and screen_height is not None:
        show_info_overlay(screen_width, screen_height)


def update_overlay_region() -> None:
    update_anchor_overlay_region()
    update_capture_overlay_region()


def destroy_all_overlays() -> None:
    hide_anchor_overlay()
    hide_capture_overlay()
    hide_info_overlay()
