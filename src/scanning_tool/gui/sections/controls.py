"""Controls section — continuous capture interval + primary action buttons."""

import tkinter as tk
from tkinter import ttk

from scanning_tool.config import save_config
from scanning_tool.gui.sections.base import SectionContext
from scanning_tool.gui.theme import style_spinbox
from scanning_tool.gui.overlays import (
    choose_label_color,
    toggle_border,
    update_overlay_region,
)
from scanning_tool.scanning import capture_once, toggle_continuous
from scanning_tool.state_context import app_state


class ControlsSection:
    """Capture-interval spinbox and the six primary action buttons."""

    def build(self, parent: ttk.Widget, ctx: SectionContext) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text="Controls", style="Glass.TLabelframe")
        frame.pack(fill="x", padx=5, pady=8)

        self._status = ctx.status

        interval_row = ttk.Frame(frame, style="Glass.Section.TFrame")
        interval_row.pack(fill="x", padx=5, pady=(5, 10))
        ttk.Label(
            interval_row, text="Continuous capture interval (s)", style="Glass.Small.TLabel"
        ).pack(side="left")

        self._interval_var = tk.DoubleVar(value=float(app_state.settings.capture.continuous_capture_interval))
        spinbox = tk.Spinbox(
            interval_row, from_=0.2, to=30.0, increment=0.1,
            textvariable=self._interval_var, width=6, format="%.1f",
            command=self._on_interval_change,
        )
        spinbox.pack(side="left", padx=5)
        style_spinbox(spinbox, ctx.colors)
        self._interval_var.trace_add("write", self._on_interval_change)

        button_row = ttk.Frame(frame, style="Glass.Section.TFrame")
        button_row.pack(fill="x", padx=5, pady=(0, 5))

        for label, command in (
            ("Single Scan", capture_once),
            ("Loop Toggle", toggle_continuous),
            ("Update Overlay", update_overlay_region),
            ("Set Label Color", choose_label_color),
            ("Save Config", save_config),
            ("Toggle Border", toggle_border),
        ):
            ttk.Button(button_row, text=label, command=command, style="Glass.TButton").pack(
                side="left", padx=5
            )

        return frame

    def _on_interval_change(self, *_args: object) -> None:
        try:
            value = float(self._interval_var.get())
        except (tk.TclError, ValueError):
            return
        value = max(0.2, min(30.0, value))
        app_state.settings.capture.continuous_capture_interval = value
        self._status.set_status(
            f"Continuous capture interval set to {app_state.settings.capture.continuous_capture_interval:.1f}s"
        )
