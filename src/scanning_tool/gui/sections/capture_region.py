"""Capture Region section — 4 sliders controlling the OCR capture rectangle via AppContext capture settings."""

from tkinter import ttk

from scanning_tool.gui.sections.base import SectionContext
from scanning_tool.gui.widgets import create_glass_scale
from scanning_tool.gui.overlays import (
    register_capture_sliders,
    sync_capture_sliders,
    update_capture_overlay_region,
)
from scanning_tool.state_context import app_state


class CaptureRegionSection:
    """Left/Top/Width/Height sliders bound to ``app_state.cap_region``."""

    def build(self, parent: ttk.Widget, ctx: SectionContext) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text="Capture Region", style="Glass.TLabelframe")
        frame.pack(fill="x", padx=5, pady=8)

        self._status = ctx.status

        cap_region = app_state.settings.capture.cap_region

        self._left = create_glass_scale(
            frame, text="Left", minimum=0, maximum=3000,
            initial=cap_region["left"], command=self._on_change,
        )
        self._top = create_glass_scale(
            frame, text="Top", minimum=0, maximum=2000,
            initial=cap_region["top"], command=self._on_change,
        )
        self._width = create_glass_scale(
            frame, text="Width", minimum=50, maximum=1000,
            initial=cap_region["width"], command=self._on_change,
        )
        self._height = create_glass_scale(
            frame, text="Height", minimum=20, maximum=500,
            initial=cap_region["height"], command=self._on_change,
            padding=(0, 0),
        )

        register_capture_sliders(self._left, self._top, self._width, self._height)
        sync_capture_sliders()
        return frame

    def _on_change(self, *_args: object) -> None:
        if app_state.control_state.gui_control_state["syncing"]["capture"]:
            return
        cap_region = app_state.settings.capture.cap_region
        cap_region["left"] = int(self._left.get())
        cap_region["top"] = int(self._top.get())
        cap_region["width"] = int(self._width.get())
        cap_region["height"] = int(self._height.get())
        self._status.set_status(f"CAP_REGION updated: {cap_region}")
        update_capture_overlay_region()
