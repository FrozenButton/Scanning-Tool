"""Ollama installation, service management, and client access."""

import logging
import os
import shutil
import socket
import subprocess
import sys
import time
import webbrowser
from typing import List, Optional, Tuple
from urllib.parse import urlparse

import ollama

from scanning_tool.state import app_state
from scanning_tool.config import save_config

logger = logging.getLogger("scanning_tool")


def sanitize_ollama_host(value: str) -> str:
    """Return a normalized Ollama host string, adding http:// when missing."""
    host = (value or "").strip()
    if not host:
        return ""
    if not app_state.host_scheme_re.match(host):
        host = f"http://{host}"
    return host


def reset_ollama_client() -> None:
    """Clear the cached Ollama client so the next call uses the latest host."""
    app_state.ollama_client = None
    app_state.ollama_client_host = ""


def set_configured_ollama_host(value: str) -> str:
    """Update the configured Ollama host and refresh environment/client state."""
    sanitized = sanitize_ollama_host(value)
    if sanitized != app_state.configured_ollama_host:
        app_state.configured_ollama_host = sanitized
        if sanitized:
            os.environ["OLLAMA_HOST"] = sanitized
        else:
            os.environ.pop("OLLAMA_HOST", None)
        reset_ollama_client()
    return sanitized


def set_configured_ollama_model(value: str) -> str:
    """Update the configured Ollama model and persist it to the config."""
    sanitized = (value or "").strip()
    if sanitized:
        if sanitized != app_state.ollama_model:
            app_state.ollama_model = sanitized
            os.environ["OLLAMA_MODEL"] = sanitized
            save_config()
    else:
        app_state.ollama_model = ""
        os.environ.pop("OLLAMA_MODEL", None)
        save_config()
    return sanitized


def get_ollama_client() -> ollama.Client:
    """Return an Ollama client instance configured for the active host."""
    host = get_ollama_host()
    if app_state.ollama_client is None or app_state.ollama_client_host != host:
        app_state.ollama_client = ollama.Client(host=host)
        app_state.ollama_client_host = host
    return app_state.ollama_client


def get_ollama_host() -> str:
    """Return the Ollama host configured via environment variable or the default."""
    env_host = os.getenv("OLLAMA_HOST", "").strip()
    if env_host:
        return sanitize_ollama_host(env_host)
    if app_state.configured_ollama_host:
        return app_state.configured_ollama_host
    return app_state.default_ollama_host


def get_ollama_model() -> str:
    """Return the active Ollama model, preferring environment config over the saved fallback."""
    env_model = os.getenv("OLLAMA_MODEL", "").strip()
    return env_model or app_state.ollama_model


def _normalize_for_parse(host: str) -> str:
    return host if "://" in host else f"http://{host}"


def is_local_ollama_host(host: str) -> bool:
    """Determine if the given host string refers to the local machine."""
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


def is_ollama_running(host: str, timeout: float = 2.0) -> bool:
    """Check whether an Ollama service is listening at the provided host."""
    hostname, port = _get_host_port(host)
    try:
        with socket.create_connection((hostname, port), timeout=timeout):
            return True
    except OSError:
        return False


def start_local_ollama_service(host: str, wait_seconds: float = 10.0) -> bool:
    """Attempt to launch `ollama serve` locally and wait for readiness."""
    if not shutil.which("ollama"):
        logger.warning("Cannot start Ollama automatically because it is not installed.")
        return False

    if app_state.ollama_server_process and app_state.ollama_server_process.poll() is None:
        return True

    logger.info("Starting local Ollama service with 'ollama serve'...")
    try:
        app_state.ollama_server_process = subprocess.Popen(
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
        if app_state.ollama_server_process.poll() is not None:
            logger.error("'ollama serve' exited before the service became ready.")
            return False
        time.sleep(0.5)

    logger.warning("Timed out waiting for Ollama service to start. Please start it manually.")
    return False


def show_installation_message(system_name: str) -> None:
    """Present final installation instructions, using a GUI prompt on Windows."""
    import tkinter as tk
    import tkinter.messagebox as messagebox

    message = (
        f"Ollama installation initiated for {system_name.title()}.\n\n"
        "After installation completes:\n"
        "1. Restart this program\n"
        "2. The first scan will download the AI model automatically\n\n"
        "Visit https://ollama.com/ for troubleshooting."
    )

    if system_name == "windows":
        temp_root = None
        try:
            temp_root = tk.Tk()
            temp_root.withdraw()
            messagebox.showinfo("Ollama Installation", message, parent=temp_root)
        except Exception as exc:
            logger.debug(f"Unable to show Windows message box: {exc}")
            logger.info(message)
        else:
            logger.info(message)
        finally:
            if temp_root is not None:
                temp_root.destroy()
    else:
        logger.info(message)


def ensure_ollama_installed() -> None:
    """Check if Ollama is installed locally when required."""
    host = get_ollama_host()
    if not is_local_ollama_host(host):
        logger.info(f"Using remote Ollama host at {host}; skipping local installation check.")
        return

    if not shutil.which("ollama"):
        import platform

        system = platform.system().lower()

        logger.info("Ollama not found on your system.")
        logger.info("Ollama is required for AI-powered code recognition.")
        logger.info("")

        if system == "windows":
            logger.info("=== Windows Installation Options ===")
            logger.info("1. Automatic download and install (Recommended)")
            logger.info("2. Manual download from website")
            logger.info("")

            download_url = "https://ollama.com/download/OllamaSetup.exe"
            logger.info("Opening the Ollama download link in your default browser...")
            logger.info(f"Download URL: {download_url}")
            try:
                opened = webbrowser.open(download_url)
                if opened:
                    logger.info("Browser opened successfully. Follow the prompts to install Ollama.")
                else:
                    logger.warning("The browser did not report success. Please open the link manually if nothing happens.")
            except Exception as e:
                logger.error(f"Unable to open browser automatically: {e}")
                logger.info("Please open the link manually to download Ollama.")

        elif system == "linux":
            logger.info("=== Linux Installation Options ===")

            distro_info = ""
            package_cmd = ""

            try:
                with open("/etc/os-release", "r") as f:
                    os_release = f.read().lower()

                if "debian" in os_release or "ubuntu" in os_release or "mint" in os_release:
                    distro_info = "Debian/Ubuntu/Mint"
                    package_cmd = "curl -fsSL https://ollama.com/install.sh | sh"
                elif "arch" in os_release or "manjaro" in os_release:
                    distro_info = "Arch/Manjaro"
                    package_cmd = "sudo pacman -S ollama"
                elif "fedora" in os_release or "rhel" in os_release or "centos" in os_release:
                    distro_info = "RedHat/Fedora/CentOS"
                    package_cmd = "curl -fsSL https://ollama.com/install.sh | sh"
                elif "gentoo" in os_release or "funtoo" in os_release:
                    distro_info = "Gentoo/Funtoo"
                    package_cmd = "sudo emerge --ask ollama"
                elif "suse" in os_release or "opensuse" in os_release:
                    distro_info = "SUSE/openSUSE"
                    package_cmd = "sudo zypper install ollama"
                else:
                    distro_info = "Unknown Linux"
                    package_cmd = "curl -fsSL https://ollama.com/install.sh | sh"
            except Exception:
                distro_info = "Unknown Linux"
                package_cmd = "curl -fsSL https://ollama.com/install.sh | sh"

            logger.info(f"Detected: {distro_info}")
            logger.info(f"Recommended command: {package_cmd}")
            logger.info("")
            logger.info("1. Run the recommended installation command")
            logger.info("2. Manual installation from website")
            logger.info("")

            choice = input("Would you like to run the installation command? (y/n): ").lower().strip()

            if choice in ["y", "yes", "1", ""]:
                logger.info(f"Running: {package_cmd}")
                logger.info("Please enter your password if prompted...")
                try:
                    result = subprocess.run(package_cmd, shell=True, check=False)
                    if result.returncode == 0:
                        logger.info("Ollama installation completed!")
                        logger.info("Please restart this program to continue.")
                    else:
                        logger.warning("Installation failed or was cancelled.")
                        logger.info("You can try installing manually from https://ollama.com/")
                except Exception as e:
                    logger.error(f"Error running installation command: {e}")
                    logger.info("Please visit https://ollama.com/ for manual installation.")
            else:
                logger.info("Opening Ollama website for manual installation...")
                webbrowser.open("https://ollama.com/")

        else:
            logger.info("=== Unsupported Operating System ===")
            logger.info("This tool currently supports Windows and Linux only.")
            logger.info("Please install Ollama manually from: https://ollama.com/")
            webbrowser.open("https://ollama.com/")

        show_installation_message(system)

        input("\nPress ENTER after installing Ollama to close this program...")
        sys.exit(0)

    else:
        try:
            version = subprocess.check_output(["ollama", "--version"], text=True).strip()
            logger.info(f"Ollama found: {version}")
        except Exception as e:
            logger.error(f"Error checking Ollama: {e}")
            sys.exit("Please install Ollama and rerun this program.")


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
            sys.exit(guidance)
        raise RuntimeError(guidance) from e

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
            sys.exit("Failed to ensure Ollama model.")
        raise RuntimeError("Failed to ensure Ollama model.") from e


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
