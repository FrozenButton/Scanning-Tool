"""Common types for GUI sections."""

from dataclasses import dataclass
from typing import Protocol

import tkinter as tk
from tkinter import ttk

from scanning_tool.gui.status import StatusBar
from scanning_tool.gui.theme import GlassPalette


@dataclass(frozen=True)
class SectionContext:
    """Shared dependencies every section needs.

    Passed to each section's ``build`` method instead of closure-captured
    globals. Keeps sections loosely coupled to the rest of the app —
    they depend on these typed handles rather than on each other.
    """

    root: tk.Tk
    colors: GlassPalette
    status: StatusBar


class Section(Protocol):
    """A UI section — one LabelFrame, self-contained widgets and callbacks."""

    def build(self, parent: ttk.Widget, ctx: SectionContext) -> ttk.Frame: ...
