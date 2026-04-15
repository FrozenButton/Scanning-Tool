"""Base GUI section framework."""

from abc import ABC, abstractmethod
from typing import Any

class BaseGuiSection(ABC):
    """Abstract base class for all GUI layout sections."""
    
    @abstractmethod
    def build(self, parent: Any) -> Any:
        """Construct the widgets for this section and attach them to parent."""
        pass
    
    @abstractmethod
    def bind_events(self) -> None:
        """Bind event handlers to the constructed widgets."""
        pass
