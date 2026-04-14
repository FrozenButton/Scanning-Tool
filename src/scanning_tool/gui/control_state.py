from dataclasses import dataclass, field
from typing import Dict, Optional, Union

import tkinter as tk
from tkinter import ttk

ScaleWidget = Union[tk.Scale, ttk.Scale]

@dataclass
class ControlState:
    gui_control_state: Dict[str, object] = field(default_factory=lambda: {
        "capture": {"left": None, "top": None, "width": None, "height": None},
        "anchor": {
            "left": None,
            "top": None,
            "width": None,
            "height": None,
            "offset_x": None,
            "offset_y": None,
        },
        "overlay": {"offset_x": None, "offset_y": None},
        "syncing": {"capture": False, "anchor": False, "overlay": False},
    })

    def get_scale_widget(self, key: str) -> Optional[ScaleWidget]:
        widget = self.gui_control_state
        for part in key.split("."):
            if isinstance(widget, dict):
                widget = widget.get(part)
            else:
                return None
        return widget if isinstance(widget, (tk.Scale, ttk.Scale)) else None
