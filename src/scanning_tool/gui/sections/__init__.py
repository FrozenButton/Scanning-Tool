"""GUI sections — each owns one LabelFrame and its widgets/callbacks."""

from scanning_tool.gui.sections.base import Section, SectionContext
from scanning_tool.gui.sections.capture_region import CaptureRegionSection
from scanning_tool.gui.sections.controls import ControlsSection
from scanning_tool.gui.sections.result_display import ResultDisplaySection

__all__ = [
    "Section",
    "SectionContext",
    "CaptureRegionSection",
    "ControlsSection",
    "ResultDisplaySection",
]
