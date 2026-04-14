"""Capture overlay window lifecycle and animation."""

import tkinter as tk
from typing import Dict

from .base import CAPTURE_ANIMATION_INTERVAL_MS, create_overlay_window, safe_tk
from .geometry import compute_capture_overlay_layout
from scanning_tool.state_context import app_state


def _apply_capture_overlay_layout(*, force: bool = False) -> None:
    overlay_state = app_state.overlay_state
    if (
        not overlay_state.capture_overlay_canvas
        or not overlay_state.capture_rect_id
        or not overlay_state.capture_overlay_root
    ):
        return

    layout = compute_capture_overlay_layout()
    overlay_width = layout["overlay_width"]
    overlay_height = layout["overlay_height"]
    left = layout["left"]
    top = layout["top"]
    padding_x = layout["padding_x"]
    padding_y = layout["padding_y"]
    cap_w = layout["cap_w"]
    cap_h = layout["cap_h"]

    last = overlay_state.capture_overlay_last_layout
    size_changed = (
        force
        or last["overlay_width"] != overlay_width
        or last["overlay_height"] != overlay_height
    )
    pos_changed = force or last["left"] != left or last["top"] != top
    rect_changed = force or last["cap_w"] != cap_w or last["cap_h"] != cap_h

    if size_changed:
        safe_tk(lambda: overlay_state.capture_overlay_canvas.config(
            width=overlay_width, height=overlay_height
        ))

    if rect_changed:
        safe_tk(lambda: overlay_state.capture_overlay_canvas.coords(
            overlay_state.capture_rect_id,
            padding_x // 2,
            padding_y,
            padding_x // 2 + cap_w,
            padding_y + cap_h,
        ))

    if size_changed or pos_changed:
        safe_tk(lambda: overlay_state.capture_overlay_root.geometry(
            f"{overlay_width}x{overlay_height}+{left}+{top}"
        ))
        safe_tk(lambda: overlay_state.capture_overlay_root.lift())

    last.update({
        "overlay_width": overlay_width,
        "overlay_height": overlay_height,
        "left": left,
        "top": top,
        "cap_w": cap_w,
        "cap_h": cap_h,
    })


def _animate_capture_overlay() -> None:
    overlay_state = app_state.overlay_state
    if (
        not overlay_state.capture_overlay_root
        or not overlay_state.capture_overlay_canvas
        or not overlay_state.capture_rect_id
        or safe_tk(overlay_state.capture_overlay_root.winfo_exists, False) is False
    ):
        overlay_state.capture_overlay_animation_job = None
        return

    try:
        _apply_capture_overlay_layout()
    except tk.TclError:
        overlay_state.capture_overlay_animation_job = None
        return

    def _schedule() -> None:
        overlay_state.capture_overlay_animation_job = overlay_state.capture_overlay_root.after(
            CAPTURE_ANIMATION_INTERVAL_MS, _animate_capture_overlay
        )

    safe_tk(_schedule)


def start_capture_overlay_animation(*, force: bool = False) -> None:
    overlay_state = app_state.overlay_state
    if (
        not overlay_state.capture_overlay_root
        or not overlay_state.capture_overlay_canvas
        or not overlay_state.capture_rect_id
    ):
        return

    _apply_capture_overlay_layout(force=force)

    if overlay_state.capture_overlay_animation_job is None:
        safe_tk(lambda: setattr(
            overlay_state,
            "capture_overlay_animation_job",
            overlay_state.capture_overlay_root.after(
                CAPTURE_ANIMATION_INTERVAL_MS,
                _animate_capture_overlay,
            ),
        ))


def stop_capture_overlay_animation() -> None:
    overlay_state = app_state.overlay_state
    if overlay_state.capture_overlay_animation_job is not None and overlay_state.capture_overlay_root:
        try:
            overlay_state.capture_overlay_root.after_cancel(overlay_state.capture_overlay_animation_job)
        except (tk.TclError, ValueError):
            pass
    overlay_state.capture_overlay_animation_job = None
    overlay_state.capture_overlay_last_layout.update({
        "overlay_width": None,
        "overlay_height": None,
        "left": None,
        "top": None,
        "cap_w": None,
        "cap_h": None,
    })


def update_capture_overlay_region() -> None:
    start_capture_overlay_animation(force=True)


def show_capture_overlay() -> None:
    overlay_state = app_state.overlay_state
    if overlay_state.capture_overlay_root and safe_tk(overlay_state.capture_overlay_root.winfo_exists, False):
        try:
            stop_capture_overlay_animation()
            overlay_state.capture_overlay_root.destroy()
        except tk.TclError:
            pass
        overlay_state.capture_overlay_canvas = None
        overlay_state.capture_rect_id = None
        overlay_state.border_canvas = None

    layout = compute_capture_overlay_layout()
    overlay_state.capture_overlay_root = create_overlay_window(
        layout["overlay_width"], layout["overlay_height"], layout["left"], layout["top"]
    )

    overlay_state.capture_overlay_canvas = tk.Canvas(
        overlay_state.capture_overlay_root,
        width=layout["overlay_width"],
        height=layout["overlay_height"],
        bg="black",
        highlightthickness=0,
    )
    overlay_state.capture_overlay_canvas.pack()
    overlay_state.border_canvas = overlay_state.capture_overlay_canvas

    overlay_state.capture_rect_id = overlay_state.capture_overlay_canvas.create_rectangle(
        layout["padding_x"] // 2,
        layout["padding_y"],
        layout["padding_x"] // 2 + layout["cap_w"],
        layout["padding_y"] + layout["cap_h"],
        outline="red",
        width=3,
        tags=("border",),
    )

    start_capture_overlay_animation(force=True)
