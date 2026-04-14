from .config.settings import CaptureSettings
from .runtime.scan_state import ScanState
from .gui.overlay_state import OverlayState
from .gui.control_state import ControlState

from dataclasses import dataclass

@dataclass
class AppContext:
    settings: CaptureSettings
    scan_state: ScanState
    overlay_state: OverlayState
    control_state: ControlState

# Temporary shim for migration
app_state = AppContext(
    settings=CaptureSettings(),
    scan_state=ScanState(),
    overlay_state=OverlayState(),
    control_state=ControlState(),
)
