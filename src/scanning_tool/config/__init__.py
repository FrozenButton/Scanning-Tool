"""Configuration loading, saving, and path resolution."""

import json
import logging
import os
import sys
from pathlib import Path

logger = logging.getLogger("scanning_tool")

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


def load_config() -> None:
    """Load configuration from config.json into app_state."""
    from scanning_tool.ollama_service import sanitize_ollama_host, reset_ollama_client

    from scanning_tool.state_context import app_state

    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                data = json.load(f)
                app_state.settings.capture.cap_region = data.get(
                    "CAP_REGION", app_state.settings.capture.cap_region
                )
                app_state.settings.overlay.label_color = data.get(
                    "label_color", app_state.settings.overlay.label_color
                )
                app_state.settings.anchor.auto_align_enabled = data.get(
                    "AUTO_ALIGN_ENABLED", app_state.settings.anchor.auto_align_enabled
                )
                app_state.settings.anchor.anchor_region = data.get(
                    "ANCHOR_REGION", app_state.settings.anchor.anchor_region
                )
                app_state.settings.anchor.anchor_offset = data.get(
                    "ANCHOR_OFFSET", app_state.settings.anchor.anchor_offset
                )
                app_state.settings.anchor.anchor_threshold = data.get(
                    "ANCHOR_THRESHOLD", app_state.settings.anchor.anchor_threshold
                )
                app_state.settings.anchor.anchor_template_dir = data.get(
                    "ANCHOR_TEMPLATE_DIR", app_state.settings.anchor.anchor_template_dir
                )
                app_state.settings.anchor.alignment_poll_interval_ms = data.get(
                    "ALIGNMENT_POLL_INTERVAL_MS", app_state.settings.anchor.alignment_poll_interval_ms
                )
                app_state.settings.capture.continuous_capture_interval = data.get(
                    "CONTINUOUS_CAPTURE_INTERVAL", app_state.settings.capture.continuous_capture_interval
                )
                app_state.settings.overlay.info_overlay_offset = data.get(
                    "INFO_OVERLAY_OFFSET", app_state.settings.overlay.info_overlay_offset
                )

                configured_host = sanitize_ollama_host(
                    data.get("OLLAMA_HOST", app_state.settings.ollama.configured_ollama_host)
                )
                configured_model = data.get("OLLAMA_MODEL", app_state.settings.ollama.ollama_model)

                env_model = os.getenv("OLLAMA_MODEL", "").strip()
                if env_model:
                    configured_model = env_model
                if configured_model != app_state.settings.ollama.ollama_model:
                    app_state.settings.ollama.ollama_model = configured_model
                if configured_host != app_state.settings.ollama.configured_ollama_host:
                    app_state.settings.ollama.configured_ollama_host = configured_host
                    if app_state.settings.ollama.configured_ollama_host:
                        os.environ["OLLAMA_HOST"] = app_state.settings.ollama.configured_ollama_host
                    reset_ollama_client()
        except (json.JSONDecodeError, OSError) as e:
            logger.warning(f"Config file invalid or empty, resetting: {e}")
            save_config()
    else:
        save_config()

    ensure_anchor_directory(app_state.settings.anchor.anchor_template_dir)
    app_state.scan_state.last_alignment_info.enabled = app_state.settings.anchor.auto_align_enabled


def save_config() -> None:
    """Persist current app_state configuration to config.json."""
    from scanning_tool.state_context import app_state

    data = {
        "CAP_REGION": app_state.settings.capture.cap_region,
        "label_color": app_state.settings.overlay.label_color,
        "AUTO_ALIGN_ENABLED": app_state.settings.anchor.auto_align_enabled,
        "ANCHOR_REGION": app_state.settings.anchor.anchor_region,
        "ANCHOR_OFFSET": app_state.settings.anchor.anchor_offset,
        "ANCHOR_THRESHOLD": app_state.settings.anchor.anchor_threshold,
        "ANCHOR_TEMPLATE_DIR": app_state.settings.anchor.anchor_template_dir,
        "ALIGNMENT_POLL_INTERVAL_MS": app_state.settings.anchor.alignment_poll_interval_ms,
        "CONTINUOUS_CAPTURE_INTERVAL": app_state.settings.capture.continuous_capture_interval,
        "INFO_OVERLAY_OFFSET": app_state.settings.overlay.info_overlay_offset,
        "OLLAMA_HOST": app_state.settings.ollama.configured_ollama_host,
        "OLLAMA_MODEL": app_state.settings.ollama.ollama_model,
    }
    with open(CONFIG_FILE, "w") as f:
        json.dump(data, f, indent=4)
        f.write("\n")
    logger.info("Config saved.")
