"""Configuration loader and saver for the scanning tool."""

import json
from loguru import logger
import os
import sys
from pathlib import Path
from typing import Dict

from pydantic import BaseModel, ConfigDict, Field, ValidationError

from scanning_tool.config.settings import AppSettings
from scanning_tool.state.context import AppContext


PROJECT_ROOT = Path(__file__).resolve().parents[3]
CONFIG_FILE = PROJECT_ROOT / "config.json"
ROCK_TYPE_FILENAME = "RockType.json"
ROCK_TYPE_FILE = PROJECT_ROOT / ROCK_TYPE_FILENAME


class ConfigData(BaseModel):
    model_config = ConfigDict(extra="ignore")

    CAP_REGION: Dict[str, int] = Field(default_factory=lambda: {
        "left": 1260,
        "top": 310,
        "width": 160,
        "height": 30,
    })
    label_color: str = "yellow"
    AUTO_ALIGN_ENABLED: bool = True
    ANCHOR_REGION: Dict[str, int] = Field(default_factory=lambda: {
        "left": 1100,
        "top": 240,
        "width": 320,
        "height": 140,
    })
    ANCHOR_OFFSET: Dict[str, int] = Field(default_factory=lambda: {"x": 36, "y": 56})
    ANCHOR_THRESHOLD: float = 0.82
    ANCHOR_TEMPLATE_DIR: str = "assets/anchor_templates"
    ALIGNMENT_POLL_INTERVAL_MS: int = 500
    CONTINUOUS_CAPTURE_INTERVAL: float = 2.0
    INFO_OVERLAY_OFFSET: Dict[str, int] = Field(default_factory=lambda: {"x": 0, "y": 0})
    OLLAMA_HOST: str = ""
    OLLAMA_MODEL: str = ""

    @classmethod
    def from_settings(cls, settings: AppSettings) -> "ConfigData":
        return cls(
            CAP_REGION=settings.capture.cap_region,
            label_color=settings.overlay.label_color,
            AUTO_ALIGN_ENABLED=settings.anchor.auto_align_enabled,
            ANCHOR_REGION=settings.anchor.anchor_region,
            ANCHOR_OFFSET=settings.anchor.anchor_offset,
            ANCHOR_THRESHOLD=settings.anchor.anchor_threshold,
            ANCHOR_TEMPLATE_DIR=settings.anchor.anchor_template_dir,
            ALIGNMENT_POLL_INTERVAL_MS=settings.anchor.alignment_poll_interval_ms,
            CONTINUOUS_CAPTURE_INTERVAL=settings.capture.continuous_capture_interval,
            INFO_OVERLAY_OFFSET=settings.overlay.info_overlay_offset,
            OLLAMA_HOST=settings.ollama.configured_ollama_host,
            OLLAMA_MODEL=settings.ollama.ollama_model,
        )

    def apply_to_settings(self, settings: AppSettings) -> None:
        settings.capture.cap_region = self.CAP_REGION
        settings.overlay.label_color = self.label_color
        settings.anchor.auto_align_enabled = self.AUTO_ALIGN_ENABLED
        settings.anchor.anchor_region = self.ANCHOR_REGION
        settings.anchor.anchor_offset = self.ANCHOR_OFFSET
        settings.anchor.anchor_threshold = self.ANCHOR_THRESHOLD
        settings.anchor.anchor_template_dir = self.ANCHOR_TEMPLATE_DIR
        settings.anchor.alignment_poll_interval_ms = self.ALIGNMENT_POLL_INTERVAL_MS
        settings.capture.continuous_capture_interval = self.CONTINUOUS_CAPTURE_INTERVAL
        settings.overlay.info_overlay_offset = self.INFO_OVERLAY_OFFSET
        settings.ollama.configured_ollama_host = self.OLLAMA_HOST
        settings.ollama.ollama_model = self.OLLAMA_MODEL


def resource_path(relative_path: str) -> str:
    """Get absolute path to a resource, works for dev and for PyInstaller bundle."""
    if hasattr(sys, "_MEIPASS"):
        return str(Path(sys._MEIPASS) / relative_path)
    return str(PROJECT_ROOT / relative_path)


def ensure_anchor_directory(path: str) -> None:
    """Ensure the directory for anchor templates exists."""
    try:
        Path(path).mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        logger.warning(f"Unable to ensure anchor template directory {path}: {exc}")


def load_config(app_context: AppContext, config_file: Path = CONFIG_FILE) -> None:
    """Load configuration from a file into the provided app context."""
    from scanning_tool.ollama import reset_ollama_client, sanitize_ollama_host

    config = ConfigData()
    if config_file.exists():
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            config = ConfigData.model_validate(data)
            env_model = os.getenv("OLLAMA_MODEL", "").strip()
            if env_model:
                config.OLLAMA_MODEL = env_model

            configured_host = sanitize_ollama_host(config.OLLAMA_HOST)
            if configured_host != app_context.settings.ollama.configured_ollama_host:
                app_context.settings.ollama.configured_ollama_host = configured_host
                if configured_host:
                    os.environ["OLLAMA_HOST"] = configured_host
                reset_ollama_client()
            config.OLLAMA_HOST = configured_host
        except (json.JSONDecodeError, OSError, ValidationError) as exc:
            logger.warning(f"Config file invalid or empty, resetting: {exc}")
            save_config(app_context, config_file)
    else:
        save_config(app_context, config_file)

    config.apply_to_settings(app_context.settings)
    ensure_anchor_directory(app_context.settings.anchor.anchor_template_dir)
    app_context.scan_state.last_alignment_info.enabled = app_context.settings.anchor.auto_align_enabled


def save_config(app_context: AppContext, config_file: Path = CONFIG_FILE) -> None:
    """Persist the provided app context configuration to disk."""
    config = ConfigData.from_settings(app_context.settings)
    with open(config_file, "w", encoding="utf-8") as f:
        json.dump(config.model_dump(), f, indent=4)
        f.write("\n")
    logger.info("Config saved.")
