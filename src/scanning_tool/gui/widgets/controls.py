"""Control widgets for the scanning tool GUI."""

from typing import Callable, Optional, Tuple

import tkinter as tk
from tkinter import ttk

from scanning_tool.gui.theme import style_spinbox


def create_section_row(parent: ttk.Widget, pady: Tuple[int, int] = (0, 5)) -> ttk.Frame:
    """Create a styled row container for section controls."""
    row = ttk.Frame(parent, style="Glass.Section.TFrame")
    row.pack(fill="x", padx=5, pady=pady)
    return row


def create_glass_scale(
    parent: ttk.Widget,
    *,
    text: str,
    minimum: float,
    maximum: float,
    initial: float,
    command: Optional[Callable[[str], None]],
    resolution: float = 1.0,
    padding: Tuple[int, int] = (0, 4),
) -> ttk.Scale:
    """Create a labeled ttk.Scale with the custom glass styling."""
    container = ttk.Frame(parent, style="Glass.Section.TFrame")
    container.pack(fill="x", padx=4, pady=padding)

    value_var = tk.DoubleVar(value=initial)

    def format_value(value: float) -> str:
        if resolution and resolution < 1.0:
            return f"{value:.2f}"
        return f"{int(round(value))}"

    label_var = tk.StringVar(value=f"{text}: {format_value(initial)}")
    ttk.Label(container, textvariable=label_var, style="Glass.Small.TLabel").pack(anchor="w", padx=2)

    def on_change(raw_value: str) -> None:
        try:
            numeric = float(raw_value)
        except (TypeError, ValueError):
            numeric = value_var.get()

        if resolution:
            snapped = round(numeric / resolution) * resolution
        else:
            snapped = numeric

        if abs(snapped - value_var.get()) > 1e-9:
            value_var.set(snapped)
            numeric = snapped
        else:
            numeric = snapped

        label_var.set(f"{text}: {format_value(numeric)}")

        if command is not None:
            if resolution and resolution < 1.0:
                command(f"{numeric:.2f}")
            else:
                command(str(int(round(numeric))))

    scale = ttk.Scale(
        container,
        from_=minimum,
        to=maximum,
        orient="horizontal",
        variable=value_var,
        command=on_change,
        style="Glass.Horizontal.TScale",
    )
    scale.pack(fill="x", padx=2, pady=(2, 0))

    def update_label(*_: object) -> None:
        value = value_var.get()
        label_var.set(f"{text}: {format_value(value)}")

    value_var.trace_add("write", update_label)

    scale._glass_container = container  # type: ignore[attr-defined]
    scale._glass_value_var = value_var  # type: ignore[attr-defined]
    scale._glass_label_var = label_var  # type: ignore[attr-defined]
    scale._glass_command = command  # type: ignore[attr-defined]
    scale._glass_resolution = resolution  # type: ignore[attr-defined]

    return scale


def create_button_row(
    parent: ttk.Widget,
    buttons: list[tuple[str, Callable[[], None]]],
    style: str = "Glass.TButton",
) -> ttk.Frame:
    """Create a row of buttons with equal spacing."""
    row = create_section_row(parent)
    for label, command in buttons:
        ttk.Button(row, text=label, command=command, style=style).pack(side="left", padx=5)
    return row
