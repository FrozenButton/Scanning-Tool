"""Anchor overlay display and region updates."""

import tkinter as tk
from typing import Dict

from .base import ANCHOR_OVERLAY_PAD, create_overlay_window, safe_tk
from .geometry import compute_anchor_overlay_geometry
from scanning_tool.state import app_state


def show_anchor_overlay() -> None:
    if not app_state.anchor_overlay_visible:
        return

    if app_state.anchor_overlay_root and safe_tk(app_state.anchor_overlay_root.winfo_exists, False):
        try:
            app_state.anchor_overlay_root.destroy()
        except tk.TclError:
            pass
        app_state.anchor_overlay_canvas = None
        app_state.anchor_rect_id = None

    geometry = compute_anchor_overlay_geometry()
    app_state.anchor_overlay_root = create_overlay_window(
        geometry["width"], geometry["height"], geometry["left"], geometry["top"]
    )
    app_state.anchor_overlay_canvas = tk.Canvas(
        app_state.anchor_overlay_root,
        width=geometry["width"],
        height=geometry["height"],
        bg="black",
        highlightthickness=0,
    )
    app_state.anchor_overlay_canvas.pack()

    app_state.anchor_rect_id = app_state.anchor_overlay_canvas.create_rectangle(
        ANCHOR_OVERLAY_PAD // 2,
        ANCHOR_OVERLAY_PAD // 2,
        ANCHOR_OVERLAY_PAD // 2 + int(app_state.anchor_region["width"]),
        ANCHOR_OVERLAY_PAD // 2 + int(app_state.anchor_region["height"]),
        outline="#00d4ff",
        width=2,
    )

    app_state.anchor_overlay_canvas.create_text(
        geometry["width"] // 2,
        5,
        text="ANCHOR REGION",
        fill="#00d4ff",
        font=("Arial", 12, "bold"),
        anchor="n",
    )


def update_anchor_overlay_region() -> None:
    if (
        not app_state.anchor_overlay_visible
        or not app_state.anchor_overlay_root
        or not app_state.anchor_overlay_canvas
        or not app_state.anchor_rect_id
    ):
        return

    geometry = compute_anchor_overlay_geometry()
    safe_tk(lambda: app_state.anchor_overlay_canvas.config(
        width=geometry["width"], height=geometry["height"]
    ))
    safe_tk(lambda: app_state.anchor_overlay_canvas.coords(
        app_state.anchor_rect_id,
        ANCHOR_OVERLAY_PAD // 2,
        ANCHOR_OVERLAY_PAD // 2,
        ANCHOR_OVERLAY_PAD // 2 + int(app_state.anchor_region["width"]),
        ANCHOR_OVERLAY_PAD // 2 + int(app_state.anchor_region["height"]),
    ))
    safe_tk(lambda: app_state.anchor_overlay_root.geometry(
        f"{geometry['width']}x{geometry['height']}+{geometry['left']}+{geometry['top']}"
    ))
    safe_tk(lambda: app_state.anchor_overlay_root.lift())


def hide_anchor_overlay() -> None:
    if app_state.anchor_overlay_root and safe_tk(app_state.anchor_overlay_root.winfo_exists, False):
        try:
            app_state.anchor_overlay_root.destroy()
        except tk.TclError:
            pass

    app_state.anchor_overlay_root = None
    app_state.anchor_overlay_canvas = None
    app_state.anchor_rect_id = None
