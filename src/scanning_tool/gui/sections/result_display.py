"""Result Display section — offset sliders for the on-screen info overlay via AppContext overlay settings."""

from tkinter import ttk

from scanning_tool.gui.sections.base import SectionContext
from scanning_tool.gui.widgets import create_glass_scale
from scanning_tool.gui.overlays import (
    register_overlay_sliders,
    reposition_info_overlay,
    sync_overlay_sliders,
)
from scanning_tool.state_context import app_state


class ResultDisplaySection:
    """Display offset X/Y sliders bound to ``app_state.info_overlay_offset``."""

    def build(self, parent: ttk.Widget, ctx: SectionContext) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text="Result Display", style="Glass.TLabelframe")
        frame.pack(fill="x", padx=5, pady=8)

        self._status = ctx.status

        overlay_offset = app_state.settings.overlay.info_overlay_offset

        self._offset_x = create_glass_scale(
            frame, text="Display offset X", minimum=-800, maximum=800,
            initial=int(overlay_offset.get("x", 0)),
            command=self._on_change,
        )
        self._offset_y = create_glass_scale(
            frame, text="Display offset Y", minimum=-600, maximum=600,
            initial=int(overlay_offset.get("y", 0)),
            command=self._on_change,
            padding=(0, 0),
        )

        register_overlay_sliders(self._offset_x, self._offset_y)
        sync_overlay_sliders()
        return frame

    def _on_change(self, *_args: object) -> None:
        if app_state.control_state.gui_control_state["syncing"].get("overlay"):
            return
        overlay_offset = app_state.settings.overlay.info_overlay_offset
        overlay_offset["x"] = int(self._offset_x.get())
        overlay_offset["y"] = int(self._offset_y.get())
        self._status.set_status(
            f"Display offset updated: x={overlay_offset['x']}, "
            f"y={overlay_offset['y']}"
        )
        reposition_info_overlay()
