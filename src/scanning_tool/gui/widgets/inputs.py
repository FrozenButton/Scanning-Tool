"""Input widgets for the scanning tool GUI."""

from typing import Callable

import tkinter as tk
from tkinter import ttk

from .controls import create_section_row
from scanning_tool.gui.theme import GlassPalette


def create_labeled_spinbox(
    parent: ttk.Widget,
    text: str,
    variable: tk.Variable,
    from_: float,
    to: float,
    increment: float,
    width: int,
    command: Callable[[object], None] | None,
    colors: GlassPalette,
) -> tk.Spinbox:
    """Create a labeled spinbox row with custom glass styling."""
    row = create_section_row(parent)
    ttk.Label(row, text=text, style="Glass.Small.TLabel").pack(side="left")

    spinbox = tk.Spinbox(
        row,
        from_=from_,
        to=to,
        increment=increment,
        textvariable=variable,
        width=width,
        command=command,
    )
    spinbox.pack(side="left", padx=5)
    from scanning_tool.gui.theme import style_spinbox
    style_spinbox(spinbox, colors)

    return spinbox


def create_labeled_entry(
    parent: ttk.Widget,
    text: str,
    variable: tk.Variable,
) -> tk.Entry:
    """Create a labeled entry row with the custom glass styling."""
    row = create_section_row(parent)
    ttk.Label(row, text=text, style="Glass.Small.TLabel").pack(side="left")

    entry = tk.Entry(row, textvariable=variable, width=40)
    entry.pack(side="left", padx=5, fill="x", expand=True)
    return entry


def create_status_label(parent: ttk.Widget, variable: tk.Variable) -> ttk.Label:
    """Create a styled status label row for section text feedback."""
    row = create_section_row(parent, pady=(0, 2))
    label = ttk.Label(row, textvariable=variable, style="Glass.Small.TLabel", justify="left")
    label.pack(fill="x", padx=5)
    return label


def create_labeled_combobox(
    parent: ttk.Widget,
    text: str,
    variable: tk.Variable,
    values: list[str],
    width: int = 40,
) -> ttk.Combobox:
    """Create a labeled combobox row with the custom glass styling."""
    row = create_section_row(parent)
    ttk.Label(row, text=text, style="Glass.Small.TLabel").pack(side="left")

    combobox = ttk.Combobox(row, textvariable=variable, values=values, width=width)
    combobox.pack(side="left", padx=5, fill="x", expand=True)
    return combobox
