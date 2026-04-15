"""Factory for building and assembling GUI components."""

import tkinter as tk
from tkinter import ttk

from scanning_tool.gui.alignment import AlignmentPoller
from scanning_tool.gui.lifecycle import register_close_handler
from scanning_tool.gui.sections import (
    CaptureRegionSection,
    ControlsSection,
    HeadSwaySection,
    OllamaSection,
    ResultDisplaySection,
    Section,
    SectionContext,
)
from scanning_tool.gui.status import StatusBar
from scanning_tool.gui.theme import apply_glass_theme
from scanning_tool.gui.widgets import ScrollableFrame
from scanning_tool.gui.overlays import show_overlay


class GUIFactory:
    """Factory for building and assembling the main GUI application."""
    
    def __init__(self):
        self._section_classes = (
            CaptureRegionSection,
            HeadSwaySection,
            OllamaSection,
            ResultDisplaySection,
            ControlsSection,
        )
    
    def create_gui(self) -> None:
        """Build and run the main Tkinter control panel."""
        root = tk.Tk()
        root.title("Star Citizen Scanner Control")
        register_close_handler(root)

        colors = apply_glass_theme(root)
        status = StatusBar(root)
        status.install_as_scanning_callback()
        ctx = SectionContext(root=root, colors=colors, status=status)

        scroll = ScrollableFrame(root, colors)
        main = scroll.inner

        for section_cls in self._section_classes:
            section_cls().build(main, ctx)

        ttk.Label(
            main, textvariable=status.status_var,
            anchor="w", justify="left", style="Glass.Status.TLabel",
        ).pack(fill="x", padx=5, pady=(8, 0))
        ttk.Label(
            main, textvariable=status.anchor_status_var,
            anchor="w", justify="left", style="Glass.Subtle.TLabel",
        ).pack(fill="x", padx=5, pady=(2, 5))

        root.update_idletasks()
        show_overlay(root.winfo_screenwidth(), root.winfo_screenheight())
        AlignmentPoller(root, status).start()
        root.mainloop()


# Create a singleton instance for backward compatibility
gui_factory = GUIFactory()
def launch_gui():
    """Backward compatible entry point."""
    gui_factory.create_gui()