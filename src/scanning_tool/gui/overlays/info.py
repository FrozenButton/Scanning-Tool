"""Floating info overlay and label management."""

import time
import tkinter as tk
from tkinter import colorchooser
from typing import Optional

from .base import create_overlay_window, safe_tk
from .geometry import compute_info_overlay_geometry
from scanning_tool.state_context import app_state


def update_overlay_label(info: dict, *, code: Optional[str] = None, raw_text: Optional[str] = None) -> None:
    overlay_state = app_state.overlay_state
    overlay_settings = app_state.settings.overlay

    message = ""
    if info:
        name = info.get("name", "")
        deposits = info.get("deposits")
        message = f"{name} x{deposits}" if deposits is not None else name

    overlay_state.overlay_text = message
    if message:
        overlay_state.last_overlay_time = time.time()
    else:
        overlay_state.last_overlay_time = 0

    if overlay_state.info_overlay_canvas and overlay_state.info_text_id:
        safe_tk(lambda: overlay_state.info_overlay_canvas.itemconfig(
            overlay_state.info_text_id,
            text=overlay_state.overlay_text,
            fill=overlay_settings.label_color,
        ))


def reposition_info_overlay() -> None:
    overlay_state = app_state.overlay_state
    overlay_settings = app_state.settings.overlay

    if (
        not overlay_state.info_overlay_root
        or not overlay_state.info_overlay_canvas
        or not overlay_state.info_text_id
    ):
        return
    if safe_tk(overlay_state.info_overlay_root.winfo_exists, False) is False:
        return

    screen_width = safe_tk(overlay_state.info_overlay_root.winfo_screenwidth, 1920) or 1920
    screen_height = safe_tk(overlay_state.info_overlay_root.winfo_screenheight, 1080) or 1080

    overlay_width, overlay_height, left, top = compute_info_overlay_geometry(screen_width, screen_height)

    safe_tk(lambda: overlay_state.info_overlay_root.geometry(f"{overlay_width}x{overlay_height}+{left}+{top}"))
    safe_tk(lambda: overlay_state.info_overlay_canvas.config(width=overlay_width, height=overlay_height))
    safe_tk(lambda: overlay_state.info_overlay_canvas.coords(overlay_state.info_text_id, overlay_width // 2, overlay_height // 2))
    safe_tk(lambda: overlay_state.info_overlay_canvas.itemconfig(overlay_state.info_text_id, width=overlay_width - 60))

    overlay_state.info_overlay_geometry.update({
        "screen_width": screen_width,
        "screen_height": screen_height,
        "width": overlay_width,
        "height": overlay_height,
    })


def start_label_timeout(window: Optional[tk.Toplevel]) -> None:
    overlay_state = app_state.overlay_state
    if overlay_state.info_overlay_canvas and overlay_state.info_text_id:
        if overlay_state.last_overlay_time and (time.time() - overlay_state.last_overlay_time > 10):
            safe_tk(lambda: overlay_state.info_overlay_canvas.itemconfig(overlay_state.info_text_id, text=""))
            overlay_state.last_overlay_time = 0

    if window and safe_tk(window.winfo_exists, False):
        safe_tk(lambda: window.after(500, lambda: start_label_timeout(window)))


def choose_label_color() -> None:
    overlay_settings = app_state.settings.overlay
    overlay_state = app_state.overlay_state

    color = colorchooser.askcolor(title="Choose Label Color")[1]
    if not color:
        return
    overlay_settings.label_color = color
    if overlay_state.info_overlay_canvas and overlay_state.info_text_id:
        safe_tk(lambda: overlay_state.info_overlay_canvas.itemconfig(overlay_state.info_text_id, fill=overlay_settings.label_color))


def toggle_border() -> None:
    overlay_state = app_state.overlay_state
    overlay_state.show_border = not overlay_state.show_border
    if overlay_state.border_canvas:
        safe_tk(lambda: overlay_state.border_canvas.itemconfig(
            "border", state="normal" if overlay_state.show_border else "hidden"
        ))


def show_info_overlay(screen_width: int, screen_height: int) -> None:
    overlay_state = app_state.overlay_state
    overlay_settings = app_state.settings.overlay

    if overlay_state.info_overlay_root and safe_tk(overlay_state.info_overlay_root.winfo_exists, False):
        try:
            overlay_state.info_overlay_root.destroy()
        except tk.TclError:
            pass
        overlay_state.info_overlay_canvas = None
        overlay_state.info_text_id = None

    overlay_width, overlay_height, left, top = compute_info_overlay_geometry(screen_width, screen_height)

    overlay_state.info_overlay_root = create_overlay_window(overlay_width, overlay_height, left, top)
    overlay_state.info_overlay_canvas = tk.Canvas(
        overlay_state.info_overlay_root,
        width=overlay_width,
        height=overlay_height,
        bg="black",
        highlightthickness=0,
    )
    overlay_state.info_overlay_canvas.pack()

    overlay_state.info_text_id = overlay_state.info_overlay_canvas.create_text(
        overlay_width // 2,
        overlay_height // 2,
        text=overlay_state.overlay_text,
        fill=overlay_settings.label_color,
        font=("Arial", 18, "bold"),
        width=overlay_width - 60,
        justify="center",
    )

    overlay_state.info_overlay_geometry.update({
        "screen_width": screen_width,
        "screen_height": screen_height,
        "width": overlay_width,
        "height": overlay_height,
    })

    start_label_timeout(overlay_state.info_overlay_root)
