"""Anchor overlay display and region updates."""

import tkinter as tk
from typing import Optional

from .base import ANCHOR_OVERLAY_PAD, create_overlay_window, safe_tk
from .geometry import compute_anchor_overlay_geometry
from scanning_tool.state_context import app_state


class AnchorOverlay:
    def __init__(self) -> None:
        self.root: Optional[tk.Toplevel] = None
        self.canvas: Optional[tk.Canvas] = None
        self.rect_id: Optional[int] = None

    def show(self) -> None:
        if not app_state.overlay_state.anchor_overlay_visible:
            return

        if self.root and safe_tk(self.root.winfo_exists, False):
            try:
                self.root.destroy()
            except tk.TclError:
                pass
            self.canvas = None
            self.rect_id = None

        geometry = compute_anchor_overlay_geometry()
        self.root = create_overlay_window(geometry["width"], geometry["height"], geometry["left"], geometry["top"])
        self.canvas = tk.Canvas(
            self.root,
            width=geometry["width"],
            height=geometry["height"],
            bg="black",
            highlightthickness=0,
        )
        self.canvas.pack()

        self.rect_id = self.canvas.create_rectangle(
            ANCHOR_OVERLAY_PAD // 2,
            ANCHOR_OVERLAY_PAD // 2,
            ANCHOR_OVERLAY_PAD // 2 + int(app_state.settings.anchor.anchor_region["width"]),
            ANCHOR_OVERLAY_PAD // 2 + int(app_state.settings.anchor.anchor_region["height"]),
            outline="#00d4ff",
            width=2,
        )

        self.canvas.create_text(
            geometry["width"] // 2,
            5,
            text="ANCHOR REGION",
            fill="#00d4ff",
            font=("Arial", 12, "bold"),
            anchor="n",
        )

        app_state.overlay_state.anchor_overlay_root = self.root
        app_state.overlay_state.anchor_overlay_canvas = self.canvas
        app_state.overlay_state.anchor_rect_id = self.rect_id

    def update_region(self) -> None:
        if (
            not app_state.overlay_state.anchor_overlay_visible
            or not self.root
            or not self.canvas
            or not self.rect_id
        ):
            return

        geometry = compute_anchor_overlay_geometry()
        safe_tk(lambda: self.canvas.config(width=geometry["width"], height=geometry["height"]))
        safe_tk(lambda: self.canvas.coords(
            self.rect_id,
            ANCHOR_OVERLAY_PAD // 2,
            ANCHOR_OVERLAY_PAD // 2,
            ANCHOR_OVERLAY_PAD // 2 + int(app_state.settings.anchor.anchor_region["width"]),
            ANCHOR_OVERLAY_PAD // 2 + int(app_state.settings.anchor.anchor_region["height"]),
        ))
        safe_tk(lambda: self.root.geometry(f"{geometry['width']}x{geometry['height']}+{geometry['left']}+{geometry['top']}"))
        safe_tk(lambda: self.root.lift())

    def hide(self) -> None:
        if self.root and safe_tk(self.root.winfo_exists, False):
            try:
                self.root.destroy()
            except tk.TclError:
                pass

        self.root = None
        self.canvas = None
        self.rect_id = None

        app_state.overlay_state.anchor_overlay_root = None
        app_state.overlay_state.anchor_overlay_canvas = None
        app_state.overlay_state.anchor_rect_id = None


_anchor_overlay = AnchorOverlay()


def show_anchor_overlay() -> None:
    _anchor_overlay.show()


def update_anchor_overlay_region() -> None:
    _anchor_overlay.update_region()


def hide_anchor_overlay() -> None:
    _anchor_overlay.hide()
