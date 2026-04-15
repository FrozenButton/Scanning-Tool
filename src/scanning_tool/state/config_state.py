"""Configuration state management."""

from dataclasses import dataclass, field
from typing import Dict, Any

@dataclass
class ConfigState:
    """Manages the lifecycle state of loaded configuration."""
    active_profile: str = "default"
    overrides: Dict[str, Any] = field(default_factory=dict)
    
    def get(self, key: str, default: Any = None) -> Any:
        return self.overrides.get(key, default)
