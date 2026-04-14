"""Capture overlay window lifecycle and animation."""

import tkinter as tk
from typing import Dict

from .base import CAPTURE_ANIMATION_INTERVAL_MS, create_overlay_window, safe_tk
from .geometry import compute_capture_overlay_layout
from scanning_tool.state import app_state


def _apply_capture_overlay_layout(*, force: bool = False) -> None:
    if (
        not app_state.capture_overlay_canvas
        or not app_state.capture_rect_id
        or not app_state.capture_overlay_root
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

    last = app_state.capture_overlay_last_layout
    size_changed = (
        force
        or last["overlay_width"] != overlay_width
        or last["overlay_height"] != overlay_height
    )
    pos_changed = force or last["left"] != left or last["top"] != top
    rect_changed = force or last["cap_w"] != cap_w or last["cap_h"] != cap_h

    if size_changed:
        safe_tk(lambda: app_state.capture_overlay_canvas.config(
            width=overlay_width, height=overlay_height
        ))

    if rect_changed:
        safe_tk(lambda: app_state.capture_overlay_canvas.coords(
            app_state.capture_rect_id,
            padding_x // 2,
            padding_y,
            padding_x // 2 + cap_w,
            padding_y + cap_h,
        ))

    if size_changed or pos_changed:
        safe_tk(lambda: app_state.capture_overlay_root.geometry(
            f"{overlay_width}x{overlay_height}+{left}+{top}"
        ))
        safe_tk(lambda: app_state.capture_overlay_root.lift())

    last.update({
        "overlay_width": overlay_width,
        "overlay_height": overlay_height,
        "left": left,
        "top": top,
        "cap_w": cap_w,
        "cap_h": cap_h,
    })


def _animate_capture_overlay() -> None:
    if (
        not app_state.capture_overlay_root
        or not app_state.capture_overlay_canvas
        or not app_state.capture_rect_id
        or safe_tk(app_state.capture_overlay_root.winfo_exists, False) is False
    ):
        app_state.capture_overlay_animation_job = None
        return

    try:
        _apply_capture_overlay_layout()
    except tk.TclError:
        app_state.capture_overlay_animation_job = None
        return

    def _schedule() -> None:
        app_state.capture_overlay_animation_job = app_state.capture_overlay_root.after(
            CAPTURE_ANIMATION_INTERVAL_MS, _animate_capture_overlay
        )

    safe_tk(_schedule)


def start_capture_overlay_animation(*, force: bool = False) -> None:
    if (
        not app_state.capture_overlay_root
        or not app_state.capture_overlay_canvas
        or not app_state.capture_rect_id
    ):
        return

    _apply_capture_overlay_layout(force=force)

    if app_state.capture_overlay_animation_job is None:
        safe_tk(lambda: setattr(
            app_state,
            "capture_overlay_animation_job",
            app_state.capture_overlay_root.after(
                CAPTURE_ANIMATION_INTERVAL_MS,
                _animate_capture_overlay,
            ),
        ))


def stop_capture_overlay_animation() -> None:
    if app_state.capture_overlay_animation_job is not None and app_state.capture_overlay_root:
        try:
            app_state.capture_overlay_root.after_cancel(app_state.capture_overlay_animation_job)
        except (tk.TclError, ValueError):
            pass
    app_state.capture_overlay_animation_job = None
    app_state.capture_overlay_last_layout.update({
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
    if app_state.capture_overlay_root and safe_tk(app_state.capture_overlay_root.winfo_exists, False):
        try:
            stop_capture_overlay_animation()
            app_state.capture_overlay_root.destroy()
        except tk.TclError:
            pass
        app_state.capture_overlay_canvas = None
        app_state.capture_rect_id = None
        app_state.border_canvas = None

    layout = compute_capture_overlay_layout()
    app_state.capture_overlay_root = create_overlay_window(
        layout["overlay_width"], layout["overlay_height"], layout["left"], layout["top"]
    )

    app_state.capture_overlay_canvas = tk.Canvas(
        app_state.capture_overlay_root,
        width=layout["overlay_width"],
        height=layout["overlay_height"],
        bg="black",
        highlightthickness=0,
    )
    app_state.capture_overlay_canvas.pack()
    app_state.border_canvas = app_state.capture_overlay_canvas

    app_state.capture_rect_id = app_state.capture_overlay_canvas.create_rectangle(
        layout["padding_x"] // 2,
        layout["padding_y"],
        layout["padding_x"] // 2 + layout["cap_w"],
        layout["padding_y"] + layout["cap_h"],
        outline="red",
        width=3,
        tags=("border",),
    )

    start_capture_overlay_animation(force=True)
