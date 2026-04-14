from dataclasses import dataclass, field
from typing import Any, Dict

@dataclass
class ControlState:
    gui_control_state: Dict[str, Any] = field(default_factory=lambda: {
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
