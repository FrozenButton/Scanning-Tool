"""Status bar primitives for the scanning tool GUI."""

import time
import tkinter as tk
from typing import Dict, Optional

from scanning_tool.state_context import app_state


class StatusBar:
    """Owns the two status StringVars and the anchor status hold/dedup logic.

    The primary ``status_var`` displays general messages. The ``anchor_status_var``
    is reserved for head-sway / alignment messages and supports a brief hold
    window so explicit user actions aren't immediately overwritten by the
    alignment poller.
    """

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.status_var = tk.StringVar(value="Ready.")
        self.anchor_status_var = tk.StringVar(value="Head sway compensation ready.")
        self._alignment_cache: Dict[str, Optional[str]] = {"message": None}
        self._anchor_hold_until: float = 0.0

    def set_status(self, message: str) -> None:
        self.status_var.set(message)

    def set_status_async(self, message: str) -> None:
        """Set status and pump idle tasks — used by background scanning callbacks."""
        self.status_var.set(message)
        self.root.update_idletasks()

    def set_anchor(self, message: str, hold: float = 1.5) -> None:
        self.anchor_status_var.set(message)
        self._anchor_hold_until = time.time() + hold
        self._alignment_cache["message"] = None

    def anchor_hold_active(self) -> bool:
        return time.time() < self._anchor_hold_until

    def push_alignment_message(self, message: str) -> None:
        """Write an alignment-loop message iff it changed and no hold is active."""
        if self.anchor_hold_active():
            return
        if message == self._alignment_cache.get("message") and self.anchor_status_var.get() == message:
            return
        self.anchor_status_var.set(message)
        self._alignment_cache["message"] = message

    def install_as_scanning_callback(self) -> None:
        """Wire scanning.py's status callback to this status bar."""
        app_state.service_state.gui_status_callback = self.set_status_async
