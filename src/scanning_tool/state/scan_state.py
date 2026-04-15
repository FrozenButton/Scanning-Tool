"""Scan state management."""

from dataclasses import dataclass, field
from typing import Optional, Any
from scanning_tool.domain.models import CaptureRegion

@dataclass
class ScanState:
    """Manages the lifecycle state of the ongoing scan process."""
    is_scanning: bool = False
    continuous_mode: bool = False
    last_result: Optional[Any] = None
    anchor_tracker: Any = field(default_factory=dict)
