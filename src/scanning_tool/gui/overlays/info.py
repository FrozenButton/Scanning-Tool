"""Floating info overlay and label management."""

import time
import tkinter as tk
from tkinter import colorchooser
from typing import Optional

from .base import create_overlay_window, safe_tk
from .geometry import compute_info_overlay_geometry
from scanning_tool.state import app_state


def update_overlay_label(info: dict, *, code: Optional[str] = None, raw_text: Optional[str] = None) -> None:
    message = ""
    if info:
        name = info.get("name", "")
        deposits = info.get("deposits")
        message = f"{name} x{deposits}" if deposits is not None else name

    app_state.overlay_text = message
    if message:
        app_state.last_overlay_time = time.time()
    else:
        app_state.last_overlay_time = 0

    if app_state.info_overlay_canvas and app_state.info_text_id:
        safe_tk(lambda: app_state.info_overlay_canvas.itemconfig(
            app_state.info_text_id,
            text=app_state.overlay_text,
            fill=app_state.label_color,
        ))


def reposition_info_overlay() -> None:
    if (
        not app_state.info_overlay_root
        or not app_state.info_overlay_canvas
        or not app_state.info_text_id
    ):
        return
    if safe_tk(app_state.info_overlay_root.winfo_exists, False) is False:
        return

    screen_width = safe_tk(app_state.info_overlay_root.winfo_screenwidth, 1920) or 1920
    screen_height = safe_tk(app_state.info_overlay_root.winfo_screenheight, 1080) or 1080

    overlay_width, overlay_height, left, top = compute_info_overlay_geometry(screen_width, screen_height)

    safe_tk(lambda: app_state.info_overlay_root.geometry(f"{overlay_width}x{overlay_height}+{left}+{top}"))
    safe_tk(lambda: app_state.info_overlay_canvas.config(width=overlay_width, height=overlay_height))
    safe_tk(lambda: app_state.info_overlay_canvas.coords(app_state.info_text_id, overlay_width // 2, overlay_height // 2))
    safe_tk(lambda: app_state.info_overlay_canvas.itemconfig(app_state.info_text_id, width=overlay_width - 60))

    app_state.info_overlay_geometry.update({
        "screen_width": screen_width,
        "screen_height": screen_height,
        "width": overlay_width,
        "height": overlay_height,
    })


def start_label_timeout(window: Optional[tk.Toplevel]) -> None:
    if app_state.info_overlay_canvas and app_state.info_text_id:
        if app_state.last_overlay_time and (time.time() - app_state.last_overlay_time > 10):
            safe_tk(lambda: app_state.info_overlay_canvas.itemconfig(app_state.info_text_id, text=""))
            app_state.last_overlay_time = 0

    if window and safe_tk(window.winfo_exists, False):
        safe_tk(lambda: window.after(500, lambda: start_label_timeout(window)))


def choose_label_color() -> None:
    color = colorchooser.askcolor(title="Choose Label Color")[1]
    if not color:
        return
    app_state.label_color = color
    if app_state.info_overlay_canvas and app_state.info_text_id:
        safe_tk(lambda: app_state.info_overlay_canvas.itemconfig(app_state.info_text_id, fill=app_state.label_color))


def toggle_border() -> None:
    app_state.show_border = not app_state.show_border
    if app_state.border_canvas:
        safe_tk(lambda: app_state.border_canvas.itemconfig(
            "border", state="normal" if app_state.show_border else "hidden"
        ))


def show_info_overlay(screen_width: int, screen_height: int) -> None:
    if app_state.info_overlay_root and safe_tk(app_state.info_overlay_root.winfo_exists, False):
        try:
            app_state.info_overlay_root.destroy()
        except tk.TclError:
            pass
        app_state.info_overlay_canvas = None
        app_state.info_text_id = None

    overlay_width, overlay_height, left, top = compute_info_overlay_geometry(screen_width, screen_height)

    app_state.info_overlay_root = create_overlay_window(overlay_width, overlay_height, left, top)
    app_state.info_overlay_canvas = tk.Canvas(
        app_state.info_overlay_root,
        width=overlay_width,
        height=overlay_height,
        bg="black",
        highlightthickness=0,
    )
    app_state.info_overlay_canvas.pack()

    app_state.info_text_id = app_state.info_overlay_canvas.create_text(
        overlay_width // 2,
        overlay_height // 2,
        text=app_state.overlay_text,
        fill=app_state.label_color,
        font=("Arial", 18, "bold"),
        width=overlay_width - 60,
        justify="center",
    )

    app_state.info_overlay_geometry.update({
        "screen_width": screen_width,
        "screen_height": screen_height,
        "width": overlay_width,
        "height": overlay_height,
    })

    start_label_timeout(app_state.info_overlay_root)
