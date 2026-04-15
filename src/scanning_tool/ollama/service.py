from loguru import logger
import shutil
import socket
import subprocess
import sys
import time

from .host import get_ollama_host, is_local_ollama_host, _get_host_port
from scanning_tool.state import app_state



def is_ollama_running(host: str, timeout: float = 2.0) -> bool:
    """Check whether an Ollama service is listening at the provided host."""
    hostname, port = _get_host_port(host)
    try:
        with socket.create_connection((hostname, port), timeout=timeout):
            return True
    except OSError:
        return False


def start_local_ollama_service(host: str, wait_seconds: float = 10.0) -> bool:
    """Launch a local Ollama service and wait for it to become ready."""
    if not shutil.which("ollama"):
        logger.warning("Cannot start Ollama automatically because it is not installed.")
        return False

    if app_state.service_state.ollama_server_process and app_state.service_state.ollama_server_process.poll() is None:
        return True

    logger.info("Starting local Ollama service with 'ollama serve'...")
    try:
        app_state.service_state.ollama_server_process = subprocess.Popen(
            ["ollama", "serve"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            stdin=subprocess.DEVNULL,
            start_new_session=True,
        )
    except Exception as exc:
        logger.error(f"Unable to start Ollama service automatically: {exc}")
        return False

    deadline = time.time() + wait_seconds
    while time.time() < deadline:
        if is_ollama_running(host):
            logger.info("Ollama service is now running.")
            return True
        if app_state.service_state.ollama_server_process.poll() is not None:
            logger.error("'ollama serve' exited before the service became ready.")
            return False
        time.sleep(0.5)

    logger.warning("Timed out waiting for Ollama service to start. Please start it manually.")
    return False


def ensure_ollama_running() -> None:
    """Start the local Ollama service if needed before contacting the API."""
    host = get_ollama_host()
    if not is_local_ollama_host(host):
        logger.info(f"Using remote Ollama host at {host}; assuming it is managed externally.")
        return

    if is_ollama_running(host):
        logger.info("Local Ollama service detected.")
        return

    if start_local_ollama_service(host):
        return

    sys.exit("Unable to reach a local Ollama service. Please start 'ollama serve' and rerun.")
