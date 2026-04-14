"""Reusable widgets for the glass-themed GUI."""

from typing import Callable, Optional, Tuple

import tkinter as tk
from tkinter import ttk

from scanning_tool.gui.theme import GlassPalette


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


def create_section_row(parent: ttk.Widget, pady: Tuple[int, int] = (0, 5)) -> ttk.Frame:
    """Create a styled row container for section controls."""
    row = ttk.Frame(parent, style="Glass.Section.TFrame")
    row.pack(fill="x", padx=5, pady=pady)
    return row


def create_labeled_spinbox(
    parent: ttk.Widget,
    text: str,
    variable: tk.Variable,
    from_: float,
    to: float,
    increment: float,
    width: int,
    command: Optional[Callable[[object], None]],
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
    colors: GlassPalette,
) -> tk.Entry:
    """Create a labeled entry row with the custom glass styling."""
    row = create_section_row(parent)
    ttk.Label(row, text=text, style="Glass.Small.TLabel").pack(side="left")

    entry = tk.Entry(row, textvariable=variable, width=40)
    entry.pack(side="left", padx=5, fill="x", expand=True)
    return entry


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


class ScrollableFrame:
    """A vertically scrollable container with a glass background.

    Children should be packed into ``self.inner``.
    """

    def __init__(self, parent: tk.Widget, colors: GlassPalette) -> None:
        self.container = ttk.Frame(parent, style="Glass.Main.TFrame")
        self.container.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(
            self.container,
            background=colors["background"],
            highlightthickness=0,
            borderwidth=0,
        )
        scrollbar = ttk.Scrollbar(self.container, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = ttk.Frame(self.canvas, style="Glass.Main.TFrame", padding=20)
        self._window_id = self.canvas.create_window((15, 15), window=self.inner, anchor="nw")

        self.inner.bind("<Configure>", self._sync_scroll_region)
        self.canvas.bind("<Configure>", self._sync_scroll_region)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", lambda _e: self._scroll_linux(-1))
        self.canvas.bind_all("<Button-5>", lambda _e: self._scroll_linux(1))

    def _sync_scroll_region(self, _event: Optional[tk.Event] = None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.canvas.itemconfigure(self._window_id, width=self.canvas.winfo_width())

    def _on_mousewheel(self, event: tk.Event) -> None:
        step = int(-1 * (event.delta / 120))
        self.canvas.yview_scroll(step, "units")

    def _scroll_linux(self, direction: int) -> None:
        self.canvas.yview_scroll(direction, "units")
