from dataclasses import dataclass

from .config.settings import AppSettings
from .runtime.scan_state import ScanState
from .runtime.service_state import ServiceState
from .gui.overlay_state import OverlayState
from .gui.control_state import ControlState

@dataclass
class AppContext:
    settings: AppSettings
    scan_state: ScanState
    overlay_state: OverlayState
    control_state: ControlState
    service_state: ServiceState

# Temporary shim for migration
app_state = AppContext(
    settings=AppSettings(),
    scan_state=ScanState(),
    overlay_state=OverlayState(),
    control_state=ControlState(),
    service_state=ServiceState(),
)
