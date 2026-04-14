"""Periodic anchor-alignment polling for the GUI."""

import tkinter as tk
from typing import Optional

from scanning_tool.anchor import perform_auto_alignment
from scanning_tool.gui.status import StatusBar
from scanning_tool.state import app_state


_IDLE_ALIGNMENT_INFO = {
    "matched": False,
    "template": None,
    "score": 0.0,
    "match_left": None,
    "match_top": None,
    "capture_left": None,
    "capture_top": None,
}


class AlignmentPoller:
    """Runs ``perform_auto_alignment`` on a Tk ``after`` cadence."""

    def __init__(self, root: tk.Tk, status: StatusBar) -> None:
        self.root = root
        self.status = status

    def start(self) -> None:
        self._tick()

    def _tick(self) -> None:
        message = self._poll()

        if message:
            self.status.push_alignment_message(message)

        try:
            interval = max(100, int(app_state.alignment_poll_interval_ms))
            self.root.after(interval, self._tick)
        except tk.TclError:
            pass

    def _poll(self) -> Optional[str]:
        if not app_state.auto_align_enabled:
            app_state.last_alignment_info.update(_IDLE_ALIGNMENT_INFO)
            return "Head sway compensation disabled."

        tracker = app_state.anchor_tracker
        if tracker is None or not getattr(tracker, "templates", None):
            app_state.last_alignment_info.update(_IDLE_ALIGNMENT_INFO)
            return "Add anchor templates to enable head sway compensation."

        match_found = perform_auto_alignment()
        info = app_state.last_alignment_info
        if info.get("matched"):
            capture_msg = f"Auto alignment adjusted CAP_REGION: {app_state.cap_region}"
            if self.status.status_var.get() != capture_msg:
                self.status.set_status(capture_msg)
            return f"Anchor locked using {info['template']} (score {info['score']:.2f})."
        if not match_found:
            return "Anchor match not found. Adjust search region or add templates."
        return None
