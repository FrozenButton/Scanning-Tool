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

        self._left = self._make_capture_scale(frame, "Left", 0, 3000, cap_region["left"])
        self._top = self._make_capture_scale(frame, "Top", 0, 2000, cap_region["top"])
        self._width = self._make_capture_scale(frame, "Width", 50, 1000, cap_region["width"])
        self._height = self._make_capture_scale(
            frame, "Height", 20, 500, cap_region["height"], padding=(0, 0)
        )

        register_capture_sliders(self._left, self._top, self._width, self._height)
        sync_capture_sliders()
        return frame

    def _make_capture_scale(
        self,
        parent: ttk.Widget,
        text: str,
        minimum: float,
        maximum: float,
        initial: float,
        padding: tuple[int, int] = (0, 4),
    ) -> ttk.Scale:
        return create_glass_scale(
            parent,
            text=text,
            minimum=minimum,
            maximum=maximum,
            initial=initial,
            command=self._on_change,
            padding=padding,
        )

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
