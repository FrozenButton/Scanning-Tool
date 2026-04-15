"""Configuration service for managing application settings."""

import json
import os
from pathlib import Path
from typing import Optional

from loguru import logger
from pydantic import BaseModel, Field, ValidationError

from scanning_tool.domain.models import (
    AutoAlignmentConfig, CaptureRegion, ContinuousCaptureConfig,
    OllamaConfig, OverlayConfig, ScanConfig
)


class ConfigData(BaseModel):
    """Pydantic model for configuration data."""
    
    capture_region: CaptureRegion = Field(default_factory=lambda: CaptureRegion(1260, 310, 160, 30))
    overlay_config: OverlayConfig = Field(default_factory=lambda: OverlayConfig({"x": 0, "y": 0}, "yellow", True))
    auto_alignment: AutoAlignmentConfig = Field(default_factory=lambda: AutoAlignmentConfig(
        True,
        500,
        CaptureRegion(1100, 240, 320, 140)
    ))
    anchor_template: CaptureRegion = Field(default_factory=lambda: CaptureRegion(1100, 240, 320, 140))
    anchor_offset: Dict[str, int] = Field(default_factory=lambda: {"x": 36, "y": 56})
    anchor_threshold: float = 0.82
    anchor_template_dir: str = "assets/anchor_templates"
    alignment_poll_interval_ms: int = 500
    continuous_capture_interval: float = 2.0
    ollama_config: OllamaConfig = Field(default_factory=lambda: OllamaConfig("", None))
    scan_config: ScanConfig = Field(default_factory=lambda: ScanConfig(0.65))


class ConfigService:
    """Service for loading, saving, and managing configuration."""
    
    def __init__(self, config_file: Path = None):
        self.config_file = config_file or Path(__file__).parent.parent.parent / "config.json"
        self._config = None
    
    def load(self) -> ConfigData:
        """Load configuration from file."""
        if not self.config_file.exists():
            logger.info("Configuration file not found, creating default.")
            self._config = ConfigData()
            self.save()
            return self._config
        
        try:
            with open(self.config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            self._config = ConfigData.model_validate(data)
            
            # Check for environment variable override
            env_model = os.getenv("OLLAMA_MODEL", "").strip()
            if env_model:
                self._config.ollama_config.model = env_model
            
            return self._config
            
        except (json.JSONDecodeError, OSError, ValidationError) as exc:
            logger.warning(f"Failed to load configuration: {exc}, using defaults.")
            self._config = ConfigData()
            self.save()
            return self._config
    
    def save(self) -> None:
        """Save configuration to file."""
        if self._config is None:
            raise ValueError("No configuration to save")
        
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self._config.model_dump(), f, indent=4)
                f.write("\n")
            logger.info("Configuration saved successfully.")
        except OSError as exc:
            logger.error(f"Failed to save configuration: {exc}")
    
    def get_config(self) -> ConfigData:
        """Get the current configuration."""
        if self._config is None:
            self.load()
        return self._config
    
    def update_config(self, **kwargs) -> None:
        """Update configuration with new values."""
        if self._config is None:
            self.load()
        
        for key, value in kwargs.items():
            if hasattr(self._config, key):
                setattr(self._config, key, value)
            else:
                logger.warning(f"Unknown config key: {key}")
    
    def get_capture_region(self) -> CaptureRegion:
        """Get the capture region configuration."""
        return self.get_config().capture_region
    
    def get_ollama_config(self) -> OllamaConfig:
        """Get the Ollama configuration."""
        return self.get_config().ollama_config
    
    def get_auto_alignment_config(self) -> AutoAlignmentConfig:
        """Get the auto-alignment configuration."""
        return self.get_config().auto_alignment
    
    def get_overlay_config(self) -> OverlayConfig:
        """Get the overlay configuration."""
        return self.get_config().overlay_config
    
    def get_scan_config(self) -> ScanConfig:
        """Get the scan configuration."""
        return self.get_config().scan_config