"""Domain models for the scanning tool configuration and state."""

from dataclasses import dataclass
from typing import Dict, Optional


@dataclass
class CaptureRegion:
    """Represents a capture region on the screen."""
    left: int
    top: int
    width: int
    height: int


@dataclass
class AnchorTemplate:
    """Represents an anchor template configuration."""
    offset: Dict[str, int]
    threshold: float
    template_dir: str


@dataclass
class OverlayConfig:
    """Represents overlay display configuration."""
    info_offset: Dict[str, int]
    label_color: str
    show_debug: bool


@dataclass
class OllamaConfig:
    """Represents Ollama AI service configuration."""
    model: str
    host: Optional[str]
    default_host: str = "http://127.0.0.1:11434"


@dataclass
class ScanConfig:
    """Represents scanning configuration."""
    min_confidence: float


@dataclass
class AutoAlignmentConfig:
    """Represents auto-alignment configuration."""
    enabled: bool
    poll_interval_ms: int
    anchor_region: CaptureRegion


@dataclass
class ContinuousCaptureConfig:
    """Represents continuous capture configuration."""
    interval: float