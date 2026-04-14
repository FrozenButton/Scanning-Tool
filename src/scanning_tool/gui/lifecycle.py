"""Window lifecycle / teardown for the GUI."""

import tkinter as tk

from scanning_tool.config import save_config
from scanning_tool.gui.overlays import stop_capture_overlay_animation
from scanning_tool.state import app_state


def register_close_handler(root: tk.Tk) -> None:
    """Wire the root window's close button to a clean teardown sequence."""

    def on_close() -> None:
        stop_capture_overlay_animation()
        save_config()
        try:
            for window in (
                app_state.capture_overlay_root,
                app_state.anchor_overlay_root,
                app_state.info_overlay_root,
            ):
                if window and window.winfo_exists():
                    window.destroy()
        except Exception:
            pass

        app_state.capture_overlay_root = None
        app_state.capture_overlay_canvas = None
        app_state.capture_rect_id = None
        app_state.anchor_overlay_root = None
        app_state.anchor_overlay_canvas = None
        app_state.anchor_rect_id = None
        app_state.info_overlay_root = None
        app_state.info_overlay_canvas = None
        app_state.info_text_id = None

        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
