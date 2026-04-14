from dataclasses import dataclass, field
from typing import Dict

@dataclass
class CaptureSettings:
    cap_region: Dict[str, int] = field(default_factory=lambda: {"left": 1260, "top": 310, "width": 160, "height": 30})
    anchor_region: Dict[str, int] = field(default_factory=lambda: {"left": 1100, "top": 240, "width": 320, "height": 140})
    anchor_offset: Dict[str, int] = field(default_factory=lambda: {"x": 36, "y": 56})
    anchor_threshold: float = 0.82
    auto_align_enabled: bool = True
    anchor_template_dir: str = "assets/anchor_templates"
    alignment_poll_interval_ms: int = 500
    continuous_capture_interval: float = 2.0
    info_overlay_offset: Dict[str, int] = field(default_factory=lambda: {"x": 0, "y": 0})
    label_color: str = "yellow"
    ollama_model: str = ""
    configured_ollama_host: str = ""
    default_ollama_host: str = "http://127.0.0.1:11434"
    min_confidence: float = 0.65
    debug_show_overlay: bool = True
