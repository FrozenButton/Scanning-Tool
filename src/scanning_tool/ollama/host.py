import os
from typing import Tuple
from urllib.parse import urlparse

from scanning_tool.config import save_config
from scanning_tool.state import app_state


def sanitize_ollama_host(value: str) -> str:
    """Return a normalized Ollama host string, adding http:// when missing."""
    host = (value or "").strip()
    if not host:
        return ""
    if not app_state.service_state.host_scheme_re.match(host):
        host = f"http://{host}"
    return host


def get_ollama_host() -> str:
    """Return the configured Ollama host, preferring environment config."""
    env_host = os.getenv("OLLAMA_HOST", "").strip()
    if env_host:
        return sanitize_ollama_host(env_host)
    if app_state.settings.ollama.configured_ollama_host:
        return app_state.settings.ollama.configured_ollama_host
    return app_state.settings.ollama.default_ollama_host


def set_configured_ollama_model(value: str) -> str:
    """Update the configured Ollama model and persist it to the config."""
    sanitized = (value or "").strip()
    if sanitized:
        if sanitized != app_state.settings.ollama.ollama_model:
            app_state.settings.ollama.ollama_model = sanitized
            os.environ["OLLAMA_MODEL"] = sanitized
            save_config(app_state)
    else:
        app_state.settings.ollama.ollama_model = ""
        os.environ.pop("OLLAMA_MODEL", None)
        save_config(app_state)
    return sanitized


def set_configured_ollama_host(value: str) -> str:
    """Update the configured Ollama host and refresh environment state."""
    sanitized = sanitize_ollama_host(value)
    if sanitized != app_state.settings.ollama.configured_ollama_host:
        app_state.settings.ollama.configured_ollama_host = sanitized
        if sanitized:
            os.environ["OLLAMA_HOST"] = sanitized
        else:
            os.environ.pop("OLLAMA_HOST", None)
    return sanitized


def get_ollama_model() -> str:
    """Return the active Ollama model, preferring environment config."""
    env_model = os.getenv("OLLAMA_MODEL", "").strip()
    return env_model or app_state.settings.ollama.ollama_model


def _normalize_for_parse(host: str) -> str:
    return host if "://" in host else f"http://{host}"


def is_local_ollama_host(host: str) -> bool:
    """Return whether the host string refers to a local Ollama host."""
    try:
        parsed = urlparse(_normalize_for_parse(host))
    except Exception:
        return True
    hostname = (parsed.hostname or "").strip().lower()
    if not hostname or hostname in {"localhost", "0.0.0.0", "127.0.0.1", "::1"}:
        return True
    if hostname.startswith("127."):
        return True
    return False


def _get_host_port(host: str) -> Tuple[str, int]:
    """Return hostname and port for the given Ollama host string."""
    parsed = urlparse(_normalize_for_parse(host))
    hostname = parsed.hostname or "127.0.0.1"
    port = parsed.port or 11434
    return hostname, port
