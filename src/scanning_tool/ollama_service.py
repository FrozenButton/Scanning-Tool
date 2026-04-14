"""Legacy compatibility layer for the old Ollama service module."""

from scanning_tool.ollama import *  # noqa: F401,F403

__all__ = [
    "get_ollama_client",
    "reset_ollama_client",
    "get_ollama_host",
    "get_ollama_model",
    "sanitize_ollama_host",
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
