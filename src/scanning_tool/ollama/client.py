import ollama
from scanning_tool.state_context import app_state

from .host import get_ollama_host


def reset_ollama_client() -> None:
    """Clear the cached Ollama client so the next call uses the latest host."""
    app_state.service_state.ollama_client = None
    app_state.service_state.ollama_client_host = ""


def get_ollama_client() -> ollama.Client:
    """Return an Ollama client instance configured for the active host."""
    host = get_ollama_host()
    if app_state.service_state.ollama_client is None or app_state.service_state.ollama_client_host != host:
        app_state.service_state.ollama_client = ollama.Client(host=host)
        app_state.service_state.ollama_client_host = host
    return app_state.service_state.ollama_client
