"""Capture overlay window lifecycle and animation."""

import tkinter as tk
from typing import Dict, Optional

from .base import CAPTURE_ANIMATION_INTERVAL_MS, create_overlay_window, safe_tk
from .geometry import compute_capture_overlay_layout
from scanning_tool.state_context import app_state


class CaptureOverlay:
    def __init__(self) -> None:
        self.root: Optional[tk.Toplevel] = None
        self.canvas: Optional[tk.Canvas] = None
        self.rect_id: Optional[int] = None
        self.border_canvas: Optional[tk.Canvas] = None
        self.animation_job: Optional[str] = None
        self.last_layout: Dict[str, Optional[int]] = {
            "overlay_width": None,
            "overlay_height": None,
            "left": None,
            "top": None,
            "cap_w": None,
            "cap_h": None,
        }

    def _apply_layout(self, *, force: bool = False) -> None:
        if not self.canvas or not self.rect_id or not self.root:
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

        last = self.last_layout
        size_changed = (
            force
            or last["overlay_width"] != overlay_width
            or last["overlay_height"] != overlay_height
        )
        pos_changed = force or last["left"] != left or last["top"] != top
        rect_changed = force or last["cap_w"] != cap_w or last["cap_h"] != cap_h

        if size_changed:
            safe_tk(lambda: self.canvas.config(width=overlay_width, height=overlay_height))

        if rect_changed:
            safe_tk(lambda: self.canvas.coords(
                self.rect_id,
                padding_x // 2,
                padding_y,
                padding_x // 2 + cap_w,
                padding_y + cap_h,
            ))

        if size_changed or pos_changed:
            safe_tk(lambda: self.root.geometry(f"{overlay_width}x{overlay_height}+{left}+{top}"))
            safe_tk(lambda: self.root.lift())

        last.update({
            "overlay_width": overlay_width,
            "overlay_height": overlay_height,
            "left": left,
            "top": top,
            "cap_w": cap_w,
            "cap_h": cap_h,
        })

    def _schedule_animation(self) -> None:
        if not self.root:
            self.animation_job = None
            return

        self.animation_job = self.root.after(
            CAPTURE_ANIMATION_INTERVAL_MS, self._animate_overlay
        )

    def _animate_overlay(self) -> None:
        if not self.root or not self.canvas or not self.rect_id or safe_tk(self.root.winfo_exists, False) is False:
            self.animation_job = None
            return

        try:
            self._apply_layout()
        except tk.TclError:
            self.animation_job = None
            return

        safe_tk(self._schedule_animation)

    def show(self) -> None:
        if self.root and safe_tk(self.root.winfo_exists, False):
            try:
                self.stop_animation()
                self.root.destroy()
            except tk.TclError:
                pass
            self.canvas = None
            self.rect_id = None
            self.border_canvas = None

        layout = compute_capture_overlay_layout()
        self.root = create_overlay_window(layout["overlay_width"], layout["overlay_height"], layout["left"], layout["top"])
        self.canvas = tk.Canvas(
            self.root,
            width=layout["overlay_width"],
            height=layout["overlay_height"],
            bg="black",
            highlightthickness=0,
        )
        self.canvas.pack()
        self.border_canvas = self.canvas

        self.rect_id = self.canvas.create_rectangle(
            layout["padding_x"] // 2,
            layout["padding_y"],
            layout["padding_x"] // 2 + layout["cap_w"],
            layout["padding_y"] + layout["cap_h"],
            outline="red",
            width=3,
            tags=("border",),
        )

        app_state.overlay_state.capture_overlay_root = self.root
        app_state.overlay_state.capture_overlay_canvas = self.canvas
        app_state.overlay_state.capture_rect_id = self.rect_id
        app_state.overlay_state.border_canvas = self.border_canvas

        self.start_animation(force=True)

    def start_animation(self, *, force: bool = False) -> None:
        if not self.root or not self.canvas or not self.rect_id:
            return

        self._apply_layout(force=force)

        if self.animation_job is None:
            self._schedule_animation()

    def stop_animation(self) -> None:
        if self.animation_job is not None and self.root:
            try:
                self.root.after_cancel(self.animation_job)
            except (tk.TclError, ValueError):
                pass
        self.animation_job = None
        self.last_layout.update({
            "overlay_width": None,
            "overlay_height": None,
            "left": None,
            "top": None,
            "cap_w": None,
            "cap_h": None,
        })

    def update_region(self) -> None:
        self.start_animation(force=True)

    def hide(self) -> None:
        if self.root and safe_tk(self.root.winfo_exists, False):
            try:
                self.stop_animation()
                self.root.destroy()
            except tk.TclError:
                pass

        self.root = None
        self.canvas = None
        self.rect_id = None
        self.border_canvas = None
        self.animation_job = None

        app_state.overlay_state.capture_overlay_root = None
        app_state.overlay_state.capture_overlay_canvas = None
        app_state.overlay_state.capture_rect_id = None
        app_state.overlay_state.border_canvas = None


_capture_overlay = CaptureOverlay()


def start_capture_overlay_animation(*, force: bool = False) -> None:
    _capture_overlay.start_animation(force=force)


def stop_capture_overlay_animation() -> None:
    _capture_overlay.stop_animation()


def update_capture_overlay_region() -> None:
    _capture_overlay.update_region()


def show_capture_overlay() -> None:
    _capture_overlay.show()


def hide_capture_overlay() -> None:
    _capture_overlay.hide()
