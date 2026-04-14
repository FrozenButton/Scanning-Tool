import logging
import sys
from typing import List, Optional

from .client import get_ollama_client
from .host import get_ollama_host, get_ollama_model, is_local_ollama_host

logger = logging.getLogger(__name__)


def ensure_model_installed(model: Optional[str] = None, exit_on_error: bool = True) -> bool:
    """Ensure the Ollama model exists on the configured host."""
    if model is None:
        model = get_ollama_model()

    host = get_ollama_host()
    host_mode = "local" if is_local_ollama_host(host) else "remote"
    logger.info(f"Using {host_mode} Ollama host at {host}.")

    client = get_ollama_client()

    try:
        response = client.list()
        available_models = {
            getattr(m, "model")
            for m in getattr(response, "models", [])
            if getattr(m, "model", None)
        }
    except Exception as e:
        logger.error(f"Unable to communicate with Ollama at {host}: {e}")
        guidance = (
            "Make sure the Ollama service is running on this PC."
            if host_mode == "local"
            else f"Ensure the Ollama server at {host} is reachable from this machine."
        )
        if exit_on_error:
            raise RuntimeError(guidance) from e
        raise

    if model in available_models:
        logger.info(f"Model {model} already available on Ollama host {host}.")
        return True

    logger.info(f"Model {model} not found on Ollama host {host}. Pulling now...")
    try:
        progress = client.pull(model)
        status = getattr(progress, "status", None)
        if status:
            logger.info(f"Ollama pull status: {status}")
        logger.info(f"Model {model} installed successfully on {host}.")
        return True
    except Exception as e:
        logger.error(f"Error ensuring model {model} on {host}: {e}")
        if exit_on_error:
            raise RuntimeError("Failed to ensure Ollama model.") from e
        raise


def list_running_ollama_models() -> List[str]:
    """Return a list of Ollama model names currently running on the active host."""
    client = get_ollama_client()
    try:
        response = client.ps()
        return [
            getattr(m, "model")
            for m in getattr(response, "models", [])
            if getattr(m, "model", None)
        ]
    except Exception as e:
        logger.warning("Unable to query Ollama process list: %s", e)
        return []


def is_model_running(model: Optional[str] = None) -> bool:
    """Return whether the given Ollama model is currently running."""
    model = model or get_ollama_model()
    return model in list_running_ollama_models()


def log_model_running_status(model: Optional[str] = None) -> bool:
    """Log the running state of the given Ollama model."""
    model = model or get_ollama_model()
    running_models = list_running_ollama_models()
    if model in running_models:
        logger.info("Ollama model %s is currently running.", model)
        return True
    logger.info(
        "Ollama model %s is not currently running. It will start on first OCR request.",
        model,
    )
    if running_models:
        logger.info("Currently running Ollama models: %s", ", ".join(running_models))
    else:
        logger.info("No Ollama models are currently running.")
    return False
