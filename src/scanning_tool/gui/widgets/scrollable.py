"""Scrollable container widget for the scanning tool GUI."""

from typing import Optional

import tkinter as tk
from tkinter import ttk

from scanning_tool.gui.theme import GlassPalette


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
