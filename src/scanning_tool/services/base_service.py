"""Base class implementation for Services."""

import logging
from abc import ABC, abstractmethod
from scanning_tool.interfaces.services import IService

class BaseService(IService, ABC):
    """Abstract base class that provides common service lifecycle management."""
    
    def __init__(self):
        self.is_running = False
        self.logger = logging.getLogger(self.__class__.__name__)
        
    def start(self) -> None:
        if self.is_running:
            self.logger.warning("Service already running.")
            return
        self.is_running = True
        self._on_start()
        
    def stop(self) -> None:
        if not self.is_running:
            return
        self.is_running = False
        self._on_stop()
        
    @abstractmethod
    def _on_start(self) -> None:
        """Internal start implementation."""
        pass
        
    @abstractmethod
    def _on_stop(self) -> None:
        """Internal stop implementation."""
        pass
