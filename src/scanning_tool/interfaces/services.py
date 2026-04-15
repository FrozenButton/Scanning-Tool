"""Shared interface definitions for the scanning tool architecture.

This module houses the base abstractions that distinct workstreams will build upon.
"""

from abc import ABC, abstractmethod
from typing import Any, Dict, Optional, Protocol


class IService(Protocol):
    """Base protocol for all background and functional services."""
    
    def start(self) -> None:
        """Starts the service."""
        ...
        
    def stop(self) -> None:
        """Stops the service."""
        ...


class IOcrService(Protocol):
    """Protocol for OCR-related capabilities."""
    
    def extract_text(self, image: Any) -> str:
        """Extract text from the given image."""
        ...

class ICaptureService(Protocol):
    """Protocol for screen capture capabilities."""
    
    def capture_region(self, region: Dict[str, int]) -> Any:
        """Capture a specific region of the screen."""
        ...
        
class IConfigService(Protocol):
    """Protocol for configuration capabilities."""
    
    def get_config(self, key: str) -> Optional[Any]:
        """Fetch a configuration value by key."""
        ...
