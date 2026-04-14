"""Configuration loader and saver for the scanning tool."""

import json
import logging
import os
import sys
from pathlib import Path

from scanning_tool.state.context import AppContext

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parents[3]
CONFIG_FILE = PROJECT_ROOT / "config.json"
ROCK_TYPE_FILENAME = "RockType.json"
ROCK_TYPE_FILE = PROJECT_ROOT / ROCK_TYPE_FILENAME


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

    if config_file.exists():
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            app_context.settings.capture.cap_region = data.get(
                "CAP_REGION", app_context.settings.capture.cap_region
            )
            app_context.settings.overlay.label_color = data.get(
                "label_color", app_context.settings.overlay.label_color
            )
            app_context.settings.anchor.auto_align_enabled = data.get(
                "AUTO_ALIGN_ENABLED", app_context.settings.anchor.auto_align_enabled
            )
            app_context.settings.anchor.anchor_region = data.get(
                "ANCHOR_REGION", app_context.settings.anchor.anchor_region
            )
            app_context.settings.anchor.anchor_offset = data.get(
                "ANCHOR_OFFSET", app_context.settings.anchor.anchor_offset
            )
            app_context.settings.anchor.anchor_threshold = data.get(
                "ANCHOR_THRESHOLD", app_context.settings.anchor.anchor_threshold
            )
            app_context.settings.anchor.anchor_template_dir = data.get(
                "ANCHOR_TEMPLATE_DIR", app_context.settings.anchor.anchor_template_dir
            )
            app_context.settings.anchor.alignment_poll_interval_ms = data.get(
                "ALIGNMENT_POLL_INTERVAL_MS", app_context.settings.anchor.alignment_poll_interval_ms
            )
            app_context.settings.capture.continuous_capture_interval = data.get(
                "CONTINUOUS_CAPTURE_INTERVAL", app_context.settings.capture.continuous_capture_interval
            )
            app_context.settings.overlay.info_overlay_offset = data.get(
                "INFO_OVERLAY_OFFSET", app_context.settings.overlay.info_overlay_offset
            )

            configured_host = sanitize_ollama_host(
                data.get("OLLAMA_HOST", app_context.settings.ollama.configured_ollama_host)
            )
            configured_model = data.get("OLLAMA_MODEL", app_context.settings.ollama.ollama_model)

            env_model = os.getenv("OLLAMA_MODEL", "").strip()
            if env_model:
                configured_model = env_model

            if configured_model != app_context.settings.ollama.ollama_model:
                app_context.settings.ollama.ollama_model = configured_model

            if configured_host != app_context.settings.ollama.configured_ollama_host:
                app_context.settings.ollama.configured_ollama_host = configured_host
                if app_context.settings.ollama.configured_ollama_host:
                    os.environ["OLLAMA_HOST"] = app_context.settings.ollama.configured_ollama_host
                reset_ollama_client()
        except (json.JSONDecodeError, OSError) as exc:
            logger.warning(f"Config file invalid or empty, resetting: {exc}")
            save_config(app_context, config_file)
    else:
        save_config(app_context, config_file)

    ensure_anchor_directory(app_context.settings.anchor.anchor_template_dir)
    app_context.scan_state.last_alignment_info.enabled = app_context.settings.anchor.auto_align_enabled


def save_config(app_context: AppContext, config_file: Path = CONFIG_FILE) -> None:
    """Persist the provided app context configuration to disk."""
    data = {
        "CAP_REGION": app_context.settings.capture.cap_region,
        "label_color": app_context.settings.overlay.label_color,
        "AUTO_ALIGN_ENABLED": app_context.settings.anchor.auto_align_enabled,
        "ANCHOR_REGION": app_context.settings.anchor.anchor_region,
        "ANCHOR_OFFSET": app_context.settings.anchor.anchor_offset,
        "ANCHOR_THRESHOLD": app_context.settings.anchor.anchor_threshold,
        "ANCHOR_TEMPLATE_DIR": app_context.settings.anchor.anchor_template_dir,
        "ALIGNMENT_POLL_INTERVAL_MS": app_context.settings.anchor.alignment_poll_interval_ms,
        "CONTINUOUS_CAPTURE_INTERVAL": app_context.settings.capture.continuous_capture_interval,
        "INFO_OVERLAY_OFFSET": app_context.settings.overlay.info_overlay_offset,
        "OLLAMA_HOST": app_context.settings.ollama.configured_ollama_host,
        "OLLAMA_MODEL": app_context.settings.ollama.ollama_model,
    }
    with open(config_file, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)
        f.write("\n")
    logger.info("Config saved.")
