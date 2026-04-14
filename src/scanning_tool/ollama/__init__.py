"""Ollama host, client, service, installer, and model utilities."""

from .client import get_ollama_client, reset_ollama_client
from .host import (
    get_ollama_host,
    get_ollama_model,
    sanitize_ollama_host,
    set_configured_ollama_host,
    set_configured_ollama_model,
    is_local_ollama_host,
)
from .installer import ensure_ollama_installed, show_installation_message
from .models import (
    ensure_model_installed,
    list_running_ollama_models,
    is_model_running,
    log_model_running_status,
)
from .service import ensure_ollama_running, is_ollama_running, start_local_ollama_service

__all__ = [
    "get_ollama_client",
    "reset_ollama_client",
    "get_ollama_host",
    "get_ollama_model",
    "sanitize_ollama_host",
    "set_configured_ollama_host",
    "set_configured_ollama_model",
    "is_local_ollama_host",
    "ensure_ollama_installed",
    "show_installation_message",
    "ensure_model_installed",
    "list_running_ollama_models",
    "is_model_running",
    "log_model_running_status",
    "ensure_ollama_running",
    "is_ollama_running",
    "start_local_ollama_service",
]
