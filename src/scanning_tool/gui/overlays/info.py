"""Floating info overlay and label management."""

import time
import tkinter as tk
from tkinter import colorchooser
from typing import Dict, Optional

from .base import create_overlay_window, safe_tk
from .capture import _capture_overlay
from .geometry import compute_info_overlay_geometry
from scanning_tool.state_context import app_state


class InfoOverlay:
    def __init__(self) -> None:
        self.root: Optional[tk.Toplevel] = None
        self.canvas: Optional[tk.Canvas] = None
        self.text_id: Optional[int] = None
        self.info_overlay_geometry: Dict[str, Optional[int]] = {
            "screen_width": None,
            "screen_height": None,
            "width": 0,
            "height": 0,
        }
        self.overlay_text: str = ""
        self.last_overlay_time: float = 0.0

    def update_label(self, info: dict, *, code: Optional[str] = None, raw_text: Optional[str] = None) -> None:
        overlay_settings = app_state.settings.overlay

        message = ""
        if info:
            name = info.get("name", "")
            deposits = info.get("deposits")
            message = f"{name} x{deposits}" if deposits is not None else name

        self.overlay_text = message
        if message:
            self.last_overlay_time = time.time()
        else:
            self.last_overlay_time = 0

        app_state.overlay_state.overlay_text = self.overlay_text
        app_state.overlay_state.last_overlay_time = self.last_overlay_time

        if self.canvas and self.text_id:
            safe_tk(lambda: self.canvas.itemconfig(
                self.text_id,
                text=self.overlay_text,
                fill=overlay_settings.label_color,
            ))

    def reposition(self) -> None:
        if not self.root or not self.canvas or not self.text_id:
            return
        if safe_tk(self.root.winfo_exists, False) is False:
            return

        screen_width = safe_tk(self.root.winfo_screenwidth, 1920) or 1920
        screen_height = safe_tk(self.root.winfo_screenheight, 1080) or 1080

        overlay_width, overlay_height, left, top = compute_info_overlay_geometry(screen_width, screen_height)

        safe_tk(lambda: self.root.geometry(f"{overlay_width}x{overlay_height}+{left}+{top}"))
        safe_tk(lambda: self.canvas.config(width=overlay_width, height=overlay_height))
        safe_tk(lambda: self.canvas.coords(self.text_id, overlay_width // 2, overlay_height // 2))
        safe_tk(lambda: self.canvas.itemconfig(self.text_id, width=overlay_width - 60))

        self.info_overlay_geometry.update({
            "screen_width": screen_width,
            "screen_height": screen_height,
            "width": overlay_width,
            "height": overlay_height,
        })
        app_state.overlay_state.info_overlay_geometry.update(self.info_overlay_geometry)

    def start_label_timeout(self) -> None:
        if self.canvas and self.text_id:
            if self.last_overlay_time and (time.time() - self.last_overlay_time > 10):
                safe_tk(lambda: self.canvas.itemconfig(self.text_id, text=""))
                self.last_overlay_time = 0
                app_state.overlay_state.last_overlay_time = 0

        if self.root and safe_tk(self.root.winfo_exists, False):
            safe_tk(lambda: self.root.after(500, self.start_label_timeout))

    def show(self, screen_width: int, screen_height: int) -> None:
        if self.root and safe_tk(self.root.winfo_exists, False):
            try:
                self.root.destroy()
            except tk.TclError:
                pass
            self.canvas = None
            self.text_id = None

        overlay_width, overlay_height, left, top = compute_info_overlay_geometry(screen_width, screen_height)

        self.root = create_overlay_window(overlay_width, overlay_height, left, top)
        self.canvas = tk.Canvas(
            self.root,
            width=overlay_width,
            height=overlay_height,
            bg="black",
            highlightthickness=0,
        )
        self.canvas.pack()

        self.text_id = self.canvas.create_text(
            overlay_width // 2,
            overlay_height // 2,
            text=self.overlay_text,
            fill=app_state.settings.overlay.label_color,
            font=("Arial", 18, "bold"),
            width=overlay_width - 60,
            justify="center",
        )

        self.info_overlay_geometry.update({
            "screen_width": screen_width,
            "screen_height": screen_height,
            "width": overlay_width,
            "height": overlay_height,
        })
        app_state.overlay_state.info_overlay_root = self.root
        app_state.overlay_state.info_overlay_canvas = self.canvas
        app_state.overlay_state.info_text_id = self.text_id
        app_state.overlay_state.info_overlay_geometry.update(self.info_overlay_geometry)

        self.start_label_timeout()

    def hide(self) -> None:
        if self.root and safe_tk(self.root.winfo_exists, False):
            try:
                self.root.destroy()
            except tk.TclError:
                pass

        self.root = None
        self.canvas = None
        self.text_id = None

        app_state.overlay_state.info_overlay_root = None
        app_state.overlay_state.info_overlay_canvas = None
        app_state.overlay_state.info_text_id = None


_info_overlay = InfoOverlay()


def update_overlay_label(info: dict, *, code: Optional[str] = None, raw_text: Optional[str] = None) -> None:
    _info_overlay.update_label(info, code=code, raw_text=raw_text)


def reposition_info_overlay() -> None:
    _info_overlay.reposition()


def start_label_timeout(window: Optional[tk.Toplevel]) -> None:
    _info_overlay.start_label_timeout()


def choose_label_color() -> None:
    overlay_settings = app_state.settings.overlay
    color = colorchooser.askcolor(title="Choose Label Color")[1]
    if not color:
        return
    overlay_settings.label_color = color
    if _info_overlay.canvas and _info_overlay.text_id:
        safe_tk(lambda: _info_overlay.canvas.itemconfig(
            _info_overlay.text_id,
            fill=overlay_settings.label_color,
        ))


def toggle_border() -> None:
    overlay_state = app_state.overlay_state
    overlay_state.show_border = not overlay_state.show_border
    if _capture_overlay.border_canvas:
        safe_tk(lambda: _capture_overlay.border_canvas.itemconfig(
            "border", state="normal" if overlay_state.show_border else "hidden"
        ))


def show_info_overlay(screen_width: int, screen_height: int) -> None:
    _info_overlay.show(screen_width, screen_height)


def hide_info_overlay() -> None:
    _info_overlay.hide()
