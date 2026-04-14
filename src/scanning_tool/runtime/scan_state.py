from dataclasses import dataclass, field
from typing import Any, Dict

@dataclass
class LastResult:
    code: Any = None
    code_raw: Any = None
    info: Any = None
    confidence: float = 0.0
    raw_text: str = ""

@dataclass
class AlignmentInfo:
    enabled: bool = True
    matched: bool = False
    template: Any = None
    score: float = 0.0
    match_left: Any = None
    match_top: Any = None
    capture_left: Any = None
    capture_top: Any = None

@dataclass
class ScanState:
    last_result: LastResult = field(default_factory=LastResult)
    last_alignment_info: AlignmentInfo = field(default_factory=AlignmentInfo)
    continuous_mode: bool = False
    show_border: bool = True
    anchor_overlay_visible: bool = True
    anchor_tracker: Any = None
