"""Anchor overlay display and region updates."""

import tkinter as tk
from typing import Dict

from .base import ANCHOR_OVERLAY_PAD, create_overlay_window, safe_tk
from .geometry import compute_anchor_overlay_geometry
from scanning_tool.state_context import app_state


def show_anchor_overlay() -> None:
    if not app_state.overlay_state.anchor_overlay_visible:
        return

    overlay_state = app_state.overlay_state

    if overlay_state.anchor_overlay_root and safe_tk(overlay_state.anchor_overlay_root.winfo_exists, False):
        try:
            overlay_state.anchor_overlay_root.destroy()
        except tk.TclError:
            pass
        overlay_state.anchor_overlay_canvas = None
        overlay_state.anchor_rect_id = None

    geometry = compute_anchor_overlay_geometry()
    overlay_state.anchor_overlay_root = create_overlay_window(
        geometry["width"], geometry["height"], geometry["left"], geometry["top"]
    )
    overlay_state.anchor_overlay_canvas = tk.Canvas(
        overlay_state.anchor_overlay_root,
        width=geometry["width"],
        height=geometry["height"],
        bg="black",
        highlightthickness=0,
    )
    overlay_state.anchor_overlay_canvas.pack()

    overlay_state.anchor_rect_id = overlay_state.anchor_overlay_canvas.create_rectangle(
        ANCHOR_OVERLAY_PAD // 2,
        ANCHOR_OVERLAY_PAD // 2,
        ANCHOR_OVERLAY_PAD // 2 + int(app_state.settings.anchor.anchor_region["width"]),
        ANCHOR_OVERLAY_PAD // 2 + int(app_state.settings.anchor.anchor_region["height"]),
        outline="#00d4ff",
        width=2,
    )

    overlay_state.anchor_overlay_canvas.create_text(
        geometry["width"] // 2,
        5,
        text="ANCHOR REGION",
        fill="#00d4ff",
        font=("Arial", 12, "bold"),
        anchor="n",
    )


def update_anchor_overlay_region() -> None:
    overlay_state = app_state.overlay_state
    if (
        not overlay_state.anchor_overlay_visible
        or not overlay_state.anchor_overlay_root
        or not overlay_state.anchor_overlay_canvas
        or not overlay_state.anchor_rect_id
    ):
        return

    geometry = compute_anchor_overlay_geometry()
    safe_tk(lambda: overlay_state.anchor_overlay_canvas.config(
        width=geometry["width"], height=geometry["height"]
    ))
    safe_tk(lambda: overlay_state.anchor_overlay_canvas.coords(
        overlay_state.anchor_rect_id,
        ANCHOR_OVERLAY_PAD // 2,
        ANCHOR_OVERLAY_PAD // 2,
        ANCHOR_OVERLAY_PAD // 2 + int(app_state.settings.anchor.anchor_region["width"]),
        ANCHOR_OVERLAY_PAD // 2 + int(app_state.settings.anchor.anchor_region["height"]),
    ))
    safe_tk(lambda: overlay_state.anchor_overlay_root.geometry(
        f"{geometry['width']}x{geometry['height']}+{geometry['left']}+{geometry['top']}"
    ))
    safe_tk(lambda: overlay_state.anchor_overlay_root.lift())


def hide_anchor_overlay() -> None:
    overlay_state = app_state.overlay_state
    if overlay_state.anchor_overlay_root and safe_tk(overlay_state.anchor_overlay_root.winfo_exists, False):
        try:
            overlay_state.anchor_overlay_root.destroy()
        except tk.TclError:
            pass

    overlay_state.anchor_overlay_root = None
    overlay_state.anchor_overlay_canvas = None
    overlay_state.anchor_rect_id = None
