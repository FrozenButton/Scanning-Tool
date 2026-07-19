import time
import re
import json
import io
import base64
import os
import sys
import socket
from urllib.parse import urlparse
from pathlib import Path
from threading import Lock, RLock, Thread
from typing import Callable, Dict, List, Optional, Tuple, Union

from PIL import Image, ImageDraw, ImageEnhance, ImageTk
import cv2
import numpy as np
import mss
import ollama
from flask import Flask, jsonify, render_template_string, request, render_template
import tkinter as tk
from tkinter import ttk, colorchooser
import keyboard  # hotkey support
import tkinter.messagebox as messagebox
import subprocess
import shutil
import webbrowser
import logging
import logging.handlers

try:
    from rapidocr_onnxruntime import RapidOCR
except ImportError:
    RapidOCR = None

# Configure logging to both console and file
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# Create formatter
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Console handler
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(formatter)

# File handler with rotation (keeps last 5 files, max 10MB each)
file_handler = logging.handlers.RotatingFileHandler(
    'scanning_tool.log', 
    maxBytes=10*1024*1024,  # 10MB
    backupCount=5
)
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(formatter)

# Add handlers to logger
logger.addHandler(console_handler)
logger.addHandler(file_handler)


ScaleWidget = Union[tk.Scale, ttk.Scale]


def apply_glass_theme(root: tk.Tk) -> Dict[str, str]:
    """Apply a holographic "glass" inspired theme to the Tkinter UI."""

    colors: Dict[str, str] = {
        "background": "#02050f",
        "panel": "#071425",
        "accent": "#67d6ff",
        "text": "#e3f6ff",
        "muted": "#7893b5",
        "button": "#10324c",
        "button_hover": "#1c4d70",
        "border": "#164b6f",
        "glow": "#36a4ff",
        "knob": "#134064",
        "knob_active": "#1f6d9c",
        "knob_outline": "#4fc3ff",
    }

    root.configure(bg=colors["background"])
    root.option_add("*Font", "{Segoe UI} 10")
    root.option_add("*Foreground", colors["text"])
    root.option_add("*TCombobox*Listbox*Background", colors["panel"])
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure("Glass.Main.TFrame", background=colors["background"])
    style.configure("Glass.Section.TFrame", background=colors["panel"])
    style.configure(
        "Glass.TLabelframe",
        background=colors["panel"],
        foreground=colors["accent"],
        borderwidth=1,
        relief="solid",
        padding=16,
    )
    try:
        style.configure(
            "Glass.TLabelframe",
            bordercolor=colors["border"],
            lightcolor=colors["border"],
            darkcolor=colors["background"],
        )
    except tk.TclError:
        pass
    style.configure(
        "Glass.TLabelframe.Label",
        background=colors["panel"],
        foreground=colors["accent"],
        font=("Segoe UI", 11, "bold"),
    )
    style.configure("Glass.TFrame", background=colors["panel"])
    style.configure("Glass.TLabel", background=colors["panel"], foreground=colors["text"])
    style.configure(
        "Glass.Small.TLabel",
        background=colors["panel"],
        foreground=colors["muted"],
        font=("Segoe UI", 9),
    )
    style.configure(
        "Glass.Status.TLabel",
        background=colors["background"],
        foreground=colors["accent"],
        font=("Segoe UI", 10, "bold"),
    )
    style.configure(
        "Glass.Subtle.TLabel",
        background=colors["background"],
        foreground=colors["muted"],
        font=("Segoe UI", 9),
    )
    style.configure(
        "Glass.TButton",
        background=colors["button"],
        foreground=colors["text"],
        borderwidth=0,
        focusthickness=3,
        focuscolor=colors["glow"],
        padding=(14, 6),
    )
    style.map(
        "Glass.TButton",
        background=[("active", colors["button_hover"]), ("pressed", colors["button_hover"])],
        foreground=[("disabled", colors["muted"])],
    )
    style.configure(
        "Glass.Active.TButton",
        background=colors["glow"],
        foreground=colors["background"],
        borderwidth=0,
        focusthickness=3,
        focuscolor=colors["accent"],
        padding=(14, 6),
    )
    style.map(
        "Glass.Active.TButton",
        background=[("active", colors["accent"]), ("pressed", colors["accent"])],
        foreground=[("disabled", colors["muted"])],
    )
    style.configure(
        "Glass.TCheckbutton",
        background=colors["panel"],
        foreground=colors["text"],
        focuscolor=colors["glow"],
    )
    style.map(
        "Glass.TCheckbutton",
        foreground=[("active", colors["accent"]), ("selected", colors["accent"])],
    )

    def make_slider_image(fill: str, outline: str) -> ImageTk.PhotoImage:
        size = 22
        radius = 8
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle(
            (1, 1, size - 2, size - 2),
            radius=radius,
            fill=fill,
            outline=outline,
            width=2,
        )
        return ImageTk.PhotoImage(img)

    slider_normal = make_slider_image(colors["knob"], colors["knob_outline"])
    slider_active = make_slider_image(colors["knob_active"], colors["accent"])
    root._glass_slider_images = (slider_normal, slider_active)  # type: ignore[attr-defined]

    try:
        style.element_create(
            "Glass.Horizontal.Scale.slider",
            "image",
            slider_normal,
            ("active", slider_active),
            ("pressed", slider_active),
        )
    except tk.TclError:
        pass

    style.layout(
        "Glass.Horizontal.TScale",
        [
            (
                "Horizontal.Scale.trough",
                {
                    "sticky": "ew",
                    "children": [("Glass.Horizontal.Scale.slider", {"side": "left", "sticky": ""})],
                },
            )
        ],
    )
    style.configure(
        "Glass.Horizontal.TScale",
        background=colors["panel"],
        troughcolor=colors["background"],
    )

    return colors


def style_spinbox(spinbox: tk.Spinbox, colors: Dict[str, str]) -> None:
    """Apply translucent styling to a Tkinter Spinbox widget."""

    try:
        spinbox.configure(
            bg=colors["panel"],
            fg=colors["text"],
            insertbackground=colors["accent"],
            disabledbackground=colors["background"],
            highlightthickness=0,
            relief="flat",
            buttonbackground=colors["button"],
        )
    except tk.TclError:
        spinbox.configure(bg=colors["panel"], fg=colors["text"])


def create_glass_scale(
    parent: ttk.Widget,
    *,
    text: str,
    minimum: float,
    maximum: float,
    initial: float,
    command: Optional[Callable[[str], None]],
    resolution: float = 1.0,
    padding: Tuple[int, int] = (0, 4),
) -> ttk.Scale:
    """Create a labeled ttk.Scale with the custom glass styling."""

    container = ttk.Frame(parent, style="Glass.Section.TFrame")
    container.pack(fill="x", padx=4, pady=padding)

    value_var = tk.DoubleVar(value=initial)

    def format_value(value: float) -> str:
        if resolution and resolution < 1.0:
            return f"{value:.2f}"
        return f"{int(round(value))}"

    label_var = tk.StringVar(value=f"{text}: {format_value(initial)}")
    ttk.Label(container, textvariable=label_var, style="Glass.Small.TLabel").pack(anchor="w", padx=2)

    def on_change(raw_value: str) -> None:
        try:
            numeric = float(raw_value)
        except (TypeError, ValueError):
            numeric = value_var.get()

        if resolution:
            snapped = round(numeric / resolution) * resolution
        else:
            snapped = numeric

        if abs(snapped - value_var.get()) > 1e-9:
            value_var.set(snapped)
            numeric = snapped
        else:
            numeric = snapped

        label_var.set(f"{text}: {format_value(numeric)}")

        if command is not None:
            if resolution and resolution < 1.0:
                command(f"{numeric:.2f}")
            else:
                command(str(int(round(numeric))))

    scale = ttk.Scale(
        container,
        from_=minimum,
        to=maximum,
        orient="horizontal",
        variable=value_var,
        command=on_change,
        style="Glass.Horizontal.TScale",
    )
    scale.pack(fill="x", padx=2, pady=(2, 0))

    def update_label(*_: object) -> None:
        value = value_var.get()
        label_var.set(f"{text}: {format_value(value)}")

    value_var.trace_add("write", update_label)

    scale._glass_container = container  # type: ignore[attr-defined]
    scale._glass_value_var = value_var  # type: ignore[attr-defined]
    scale._glass_label_var = label_var  # type: ignore[attr-defined]
    scale._glass_command = command  # type: ignore[attr-defined]
    scale._glass_resolution = resolution  # type: ignore[attr-defined]

    return scale

def show_installation_message(system_name: str) -> None:
    """Present final installation instructions, using a GUI prompt on Windows."""
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

def ensure_ollama_installed():
    """
    Check if Ollama is installed locally when required.
    If a remote host is configured, skip the local installation prompts.
    """

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
            # Windows - offer automatic download and install
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
            # Linux - detect distribution and offer package manager commands
            logger.info("=== Linux Installation Options ===")

            # Try to detect Linux distribution
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
            except:
                distro_info = "Unknown Linux"
                package_cmd = "curl -fsSL https://ollama.com/install.sh | sh"
            
            logger.info(f"Detected: {distro_info}")
            logger.info(f"Recommended command: {package_cmd}")
            logger.info("")
            logger.info("1. Run the recommended installation command")
            logger.info("2. Manual installation from website")
            logger.info("")

            choice = input("Would you like to run the installation command? (y/n): ").lower().strip()
            
            if choice in ['y', 'yes', '1', '']:
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
            # Unsupported OS
            logger.info("=== Unsupported Operating System ===")
            logger.info("This tool currently supports Windows and Linux only.")
            logger.info("Please install Ollama manually from: https://ollama.com/")
            webbrowser.open("https://ollama.com/")
        
        # Show final message
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

def ensure_model_installed(model="qwen2.5vl:3b"):
    """Ensure the Ollama model exists on the configured host."""

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
        sys.exit(guidance)

    if model in available_models:
        logger.info(f"Model {model} already available on Ollama host {host}.")
        return

    logger.info(f"Model {model} not found on Ollama host {host}. Pulling now...")
    try:
        progress = client.pull(model)
        status = getattr(progress, "status", None)
        if status:
            logger.info(f"Ollama pull status: {status}")
        logger.info(f"Model {model} installed successfully on {host}.")
    except Exception as e:
        logger.error(f"Error ensuring model {model} on {host}: {e}")
        sys.exit("Failed to ensure Ollama model.")



# ---------- CONFIG ----------
CONFIG_FILE = "config.json"
USER_ROCK_DATA_FILE = "user_rock_data.json"
STARMINERS_ORE_DATA_FILE = "starminers_ore_data.json"

CAP_REGION = {"left": 1260, "top": 310, "width": 160, "height": 30}
ANCHOR_REGION = {"left": 1100, "top": 240, "width": 320, "height": 140}
ANCHOR_OFFSET = {"x": 36, "y": 56}
ANCHOR_THRESHOLD = 0.82
AUTO_ALIGN_ENABLED = True
ANCHOR_TEMPLATE_DIR = "assets/anchor_templates"
ALIGNMENT_POLL_INTERVAL_MS = 500
CONTINUOUS_CAPTURE_INTERVAL = 2.0
OLLAMA_KEEP_ALIVE = "30m"
OLLAMA_NUM_PREDICT = 32
OLLAMA_REQUEST_TIMEOUT = 30.0
OLLAMA_LOAD_TIMEOUT = 90.0
OLLAMA_FALLBACK_ENABLED = False
FAST_OCR_ENABLED = True
FAST_OCR_MIN_CONFIDENCE = 0.65
INFO_OVERLAY_OFFSET = {"x": 0, "y": 0}
MATCH_BOX_ENABLED = False
MATCH_BOX_OFFSET = {"x": 0, "y": 160}
MATCH_BOX_SCALE = 1.0
MATCH_BOX_OPACITY = 0.85
label_color = "yellow"
MIN_CONFIDENCE = 0.65
DEBUG_SHOW_OVERLAY = True
OLLAMA_MODEL = "qwen2.5vl:3b"   # vision model
DEFAULT_OLLAMA_HOST = "http://127.0.0.1:11434"
CONFIGURED_OLLAMA_HOST = ""

_OLLAMA_CLIENT = None
_OLLAMA_CLIENT_HOST = ""
_OLLAMA_SERVER_PROCESS: Optional[subprocess.Popen] = None
_FAST_OCR_ENGINE = None
SCAN_LOCK = Lock()
DATA_LOCK = RLock()
scan_in_progress = False
continuous_scan_thread: Optional[Thread] = None
last_scan_skip_log_time = 0.0
last_scan_started_at = 0.0
SCAN_SKIP_LOG_INTERVAL = 5.0

_HOST_SCHEME_RE = re.compile(r"^[a-zA-Z][a-zA-Z0-9+.-]*://")


def sanitize_ollama_host(value: str) -> str:
    """Return a normalized Ollama host string, adding http:// when missing."""

    host = (value or "").strip()
    if not host:
        return ""
    if not _HOST_SCHEME_RE.match(host):
        host = f"http://{host}"
    return host


def reset_ollama_client() -> None:
    """Clear the cached Ollama client so the next call uses the latest host."""

    global _OLLAMA_CLIENT, _OLLAMA_CLIENT_HOST
    _OLLAMA_CLIENT = None
    _OLLAMA_CLIENT_HOST = ""


def set_configured_ollama_host(value: str) -> str:
    """Update the configured Ollama host and refresh environment/client state."""

    global CONFIGURED_OLLAMA_HOST
    sanitized = sanitize_ollama_host(value)
    if sanitized != CONFIGURED_OLLAMA_HOST:
        CONFIGURED_OLLAMA_HOST = sanitized
        if sanitized:
            os.environ["OLLAMA_HOST"] = sanitized
        else:
            os.environ.pop("OLLAMA_HOST", None)
        reset_ollama_client()
    return sanitized


def get_ollama_client() -> "ollama.Client":
    """Return an Ollama client instance configured for the active host."""

    global _OLLAMA_CLIENT, _OLLAMA_CLIENT_HOST
    host = get_ollama_host()
    if _OLLAMA_CLIENT is None or _OLLAMA_CLIENT_HOST != host:
        _OLLAMA_CLIENT = ollama.Client(host=host, timeout=float(OLLAMA_REQUEST_TIMEOUT))
        _OLLAMA_CLIENT_HOST = host
    return _OLLAMA_CLIENT


def get_fast_ocr_engine():
    global _FAST_OCR_ENGINE
    if not FAST_OCR_ENABLED or RapidOCR is None:
        return None
    if _FAST_OCR_ENGINE is None:
        _FAST_OCR_ENGINE = RapidOCR()
    return _FAST_OCR_ENGINE


def warm_fast_ocr_engine() -> None:
    if not FAST_OCR_ENABLED:
        return
    if RapidOCR is None:
        logger.info("Fast OCR unavailable; install rapidocr-onnxruntime for sub-second scans.")
        return
    try:
        start = time.perf_counter()
        img = Image.new("RGB", (120, 44), (0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.text((8, 10), "17,080", fill=(255, 255, 255))
        text, conf = ocr_with_fast_engine(img)
        logger.info("Fast OCR warmed in %.2fs%s", time.perf_counter() - start, f": {text} ({conf:.2f})" if text else "")
    except Exception as exc:
        logger.warning("Fast OCR warm-up did not complete: %s", exc)


def warm_ollama_model(model: str = OLLAMA_MODEL) -> None:
    """Load the model once at startup so the first scan does not hit the OCR timeout."""

    client = ollama.Client(host=get_ollama_host(), timeout=max(float(OLLAMA_LOAD_TIMEOUT), float(OLLAMA_REQUEST_TIMEOUT)))
    try:
        start = time.perf_counter()
        response = client.generate(
            model=model,
            prompt="Return only: OK",
            options={"temperature": 0, "num_predict": 4, "num_ctx": 512},
            keep_alive=OLLAMA_KEEP_ALIVE,
        )
        text = (getattr(response, "response", None) or response.get("response", "")).strip()
        logger.info("Model %s warmed in %.2fs%s", model, time.perf_counter() - start, f": {text}" if text else "")
    except Exception as exc:
        logger.warning("Model warm-up did not complete: %s", exc)


def get_ollama_host() -> str:
    """Return the Ollama host configured via environment variable or the default."""

    env_host = os.getenv("OLLAMA_HOST", "").strip()
    if env_host:
        return sanitize_ollama_host(env_host)

    if CONFIGURED_OLLAMA_HOST:
        return CONFIGURED_OLLAMA_HOST

    return DEFAULT_OLLAMA_HOST


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

    global _OLLAMA_SERVER_PROCESS

    if not shutil.which("ollama"):
        logger.warning("Cannot start Ollama automatically because it is not installed.")
        return False

    if _OLLAMA_SERVER_PROCESS and _OLLAMA_SERVER_PROCESS.poll() is None:
        return True

    logger.info("Starting local Ollama service with 'ollama serve'...")
    try:
        _OLLAMA_SERVER_PROCESS = subprocess.Popen(
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
        if _OLLAMA_SERVER_PROCESS.poll() is not None:
            logger.error("'ollama serve' exited before the service became ready.")
            return False
        time.sleep(0.5)

    logger.warning("Timed out waiting for Ollama service to start. Please start it manually.")
    return False

# Regex for codes
CODE_RE = re.compile(
    r"(?:[A-Za-z]?-?\d[\d,\.]{1,10}|\d{2,10})",
    re.IGNORECASE
)

last_result = {"code": None, "code_raw": None, "ore": None,
               "quality": None, "confidence": 0.0, "raw_text": "",
               "near_ore_matches": [], "ocr_error": None}
last_match_result: Optional[Dict] = None
last_ocr_error: Optional[str] = None
last_alignment_info = {
    "enabled": AUTO_ALIGN_ENABLED,
    "matched": False,
    "template": None,
    "score": 0.0,
    "match_left": None,
    "match_top": None,
    "capture_left": None,
    "capture_top": None,
}


GUI_CONTROL_STATE = {
    "capture": {"left": None, "top": None, "width": None, "height": None},
    "anchor": {"left": None, "top": None, "width": None, "height": None, "offset_x": None, "offset_y": None},
    "overlay": {"offset_x": None, "offset_y": None},
    "syncing": {"capture": False, "anchor": False, "overlay": False},
}
GUI_TOGGLE_BUTTONS = {"loop": None, "border": None, "keys": None, "match_box": None}


def refresh_toggle_buttons() -> None:
    loop_button = GUI_TOGGLE_BUTTONS.get("loop")
    border_button = GUI_TOGGLE_BUTTONS.get("border")
    keys_button = GUI_TOGGLE_BUTTONS.get("keys")
    match_box_button = GUI_TOGGLE_BUTTONS.get("match_box")
    if loop_button is not None:
        try:
            loop_button.configure(
                text=f"Loop: {'ON' if continuous_mode else 'OFF'}",
                style="Glass.Active.TButton" if continuous_mode else "Glass.TButton",
            )
        except tk.TclError:
            pass
    if border_button is not None:
        try:
            border_button.configure(
                text=f"Border: {'ON' if show_border else 'OFF'}",
                style="Glass.Active.TButton" if show_border else "Glass.TButton",
            )
        except tk.TclError:
            pass
    if keys_button is not None:
        try:
            keys_button.configure(
                text=f"Keys: {'ON' if show_keybind_overlay else 'OFF'}",
                style="Glass.Active.TButton" if show_keybind_overlay else "Glass.TButton",
            )
        except tk.TclError:
            pass
    if match_box_button is not None:
        try:
            match_box_button.configure(
                text=f"Match Box: {'ON' if MATCH_BOX_ENABLED else 'OFF'}",
                style="Glass.Active.TButton" if MATCH_BOX_ENABLED else "Glass.TButton",
            )
        except tk.TclError:
            pass


def register_capture_sliders(left: ScaleWidget, top: ScaleWidget, width: ScaleWidget, height: ScaleWidget) -> None:
    GUI_CONTROL_STATE["capture"].update({
        "left": left,
        "top": top,
        "width": width,
        "height": height,
    })


def register_anchor_sliders(
    left: ScaleWidget,
    top: ScaleWidget,
    width: ScaleWidget,
    height: ScaleWidget,
    offset_x: ScaleWidget,
    offset_y: ScaleWidget,
) -> None:
    GUI_CONTROL_STATE["anchor"].update({
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "offset_x": offset_x,
        "offset_y": offset_y,
    })


def register_overlay_sliders(offset_x: ScaleWidget, offset_y: ScaleWidget) -> None:
    GUI_CONTROL_STATE["overlay"].update({"offset_x": offset_x, "offset_y": offset_y})


def sync_capture_sliders() -> None:
    state = GUI_CONTROL_STATE
    widgets = state["capture"]
    widget = widgets["left"]
    if not widget:
        return
    if state["syncing"]["capture"]:
        return

    def _apply() -> None:
        if state["syncing"]["capture"]:
            return
        state["syncing"]["capture"] = True
        try:
            try:
                widgets["left"].set(int(CAP_REGION["left"]))
                widgets["top"].set(int(CAP_REGION["top"]))
                widgets["width"].set(int(CAP_REGION["width"]))
                widgets["height"].set(int(CAP_REGION["height"]))
            except tk.TclError:
                pass
        finally:
            state["syncing"]["capture"] = False

    try:
        widget.after(0, _apply)
    except tk.TclError:
        pass


def sync_anchor_sliders() -> None:
    state = GUI_CONTROL_STATE
    widgets = state["anchor"]
    widget = widgets["left"]
    if not widget:
        return
    if state["syncing"]["anchor"]:
        return

    def _apply() -> None:
        if state["syncing"]["anchor"]:
            return
        state["syncing"]["anchor"] = True
        try:
            try:
                widgets["left"].set(int(ANCHOR_REGION["left"]))
                widgets["top"].set(int(ANCHOR_REGION["top"]))
                widgets["width"].set(int(ANCHOR_REGION["width"]))
                widgets["height"].set(int(ANCHOR_REGION["height"]))
                widgets["offset_x"].set(int(ANCHOR_OFFSET["x"]))
                widgets["offset_y"].set(int(ANCHOR_OFFSET["y"]))
            except tk.TclError:
                pass
        finally:
            state["syncing"]["anchor"] = False


def sync_overlay_sliders() -> None:
    state = GUI_CONTROL_STATE
    widgets = state["overlay"]
    widget = widgets["offset_x"]
    if not widget:
        return
    if state["syncing"]["overlay"]:
        return

    def _apply() -> None:
        if state["syncing"]["overlay"]:
            return
        state["syncing"]["overlay"] = True
        try:
            try:
                widgets["offset_x"].set(int(INFO_OVERLAY_OFFSET["x"]))
                widgets["offset_y"].set(int(INFO_OVERLAY_OFFSET["y"]))
            except tk.TclError:
                pass
        finally:
            state["syncing"]["overlay"] = False

    try:
        widget.after(0, _apply)
    except tk.TclError:
        pass


# ---------- Config Handling ----------
def ensure_anchor_directory(path: str) -> None:
    """Ensure the directory for anchor templates exists."""
    try:
        Path(path).mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        logger.warning(f"Unable to ensure anchor template directory {path}: {exc}")


class AnchorRegionTracker:
    """Manage template loading and anchor matching for auto alignment."""

    def __init__(self, template_dir: str, threshold: float = 0.82) -> None:
        self.template_dir = template_dir
        self.threshold = threshold
        self.templates: List[Tuple[str, np.ndarray]] = []
        self.last_loaded_count = 0
        self.load_templates()

    def set_threshold(self, threshold: float) -> None:
        self.threshold = threshold

    def set_directory(self, template_dir: str) -> int:
        self.template_dir = template_dir
        return self.load_templates()

    def load_templates(self) -> int:
        ensure_anchor_directory(self.template_dir)
        loaded: List[Tuple[str, np.ndarray]] = []
        directory = Path(self.template_dir)
        if not directory.exists():
            logger.debug(f"Anchor template directory does not exist: {directory}")
            self.templates = []
            self.last_loaded_count = 0
            return 0

        supported_ext = {".png", ".jpg", ".jpeg", ".bmp"}
        for path in sorted(directory.glob("**/*")):
            if path.suffix.lower() not in supported_ext or not path.is_file():
                continue
            image = cv2.imread(str(path), cv2.IMREAD_GRAYSCALE)
            if image is None:
                logger.warning(f"Failed to load anchor template: {path}")
                continue
            loaded.append((path.name, image))

        self.templates = loaded
        self.last_loaded_count = len(loaded)
        if self.last_loaded_count == 0:
            logger.warning(
                "No anchor templates were loaded. Head sway compensation will remain disabled until templates are added."
            )
        else:
            logger.info(f"Loaded {self.last_loaded_count} anchor templates from {directory}")
        return self.last_loaded_count

    def locate_anchor(self, region: Dict[str, int]) -> Optional[Dict[str, float]]:
        if not self.templates:
            return None

        monitor = {
            "left": int(region["left"]),
            "top": int(region["top"]),
            "width": int(region["width"]),
            "height": int(region["height"]),
        }

        with mss.mss() as sct:
            try:
                screenshot = sct.grab(monitor)
            except Exception as exc:
                logger.error(f"Anchor capture failed: {exc}")
                return None

        anchor_image = np.array(screenshot)
        if anchor_image.ndim == 3 and anchor_image.shape[2] == 4:
            anchor_gray = cv2.cvtColor(anchor_image, cv2.COLOR_BGRA2GRAY)
        else:
            anchor_gray = cv2.cvtColor(anchor_image, cv2.COLOR_BGR2GRAY)

        best_score = -1.0
        best_loc: Optional[Tuple[int, int]] = None
        best_template: Optional[Tuple[str, np.ndarray]] = None

        for template_name, template_img in self.templates:
            if anchor_gray.shape[0] < template_img.shape[0] or anchor_gray.shape[1] < template_img.shape[1]:
                continue
            res = cv2.matchTemplate(anchor_gray, template_img, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(res)
            if max_val > best_score:
                best_score = float(max_val)
                best_loc = max_loc
                best_template = (template_name, template_img)

        if best_loc is None or best_template is None:
            return None

        if best_score < self.threshold:
            logger.debug(
                f"Anchor match below threshold ({best_score:.3f} < {self.threshold:.3f}) using template {best_template[0]}"
            )
            return None

        match_left = monitor["left"] + best_loc[0]
        match_top = monitor["top"] + best_loc[1]
        return {
            "match_left": float(match_left),
            "match_top": float(match_top),
            "score": best_score,
            "template": best_template[0],
            "template_width": float(best_template[1].shape[1]),
            "template_height": float(best_template[1].shape[0]),
        }


anchor_tracker: Optional[AnchorRegionTracker] = None


def load_config():
    global CAP_REGION, label_color, AUTO_ALIGN_ENABLED, ANCHOR_REGION, ANCHOR_OFFSET, ANCHOR_THRESHOLD, ANCHOR_TEMPLATE_DIR
    global ALIGNMENT_POLL_INTERVAL_MS, CONTINUOUS_CAPTURE_INTERVAL, INFO_OVERLAY_OFFSET, MATCH_BOX_ENABLED, MATCH_BOX_OFFSET, MATCH_BOX_SCALE, MATCH_BOX_OPACITY, CONFIGURED_OLLAMA_HOST
    global OLLAMA_KEEP_ALIVE, OLLAMA_NUM_PREDICT, OLLAMA_REQUEST_TIMEOUT, OLLAMA_FALLBACK_ENABLED, FAST_OCR_ENABLED, FAST_OCR_MIN_CONFIDENCE
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                data = json.load(f)
                CAP_REGION = data.get("CAP_REGION", CAP_REGION)
                label_color = data.get("label_color", label_color)
                AUTO_ALIGN_ENABLED = data.get("AUTO_ALIGN_ENABLED", AUTO_ALIGN_ENABLED)
                ANCHOR_REGION = data.get("ANCHOR_REGION", ANCHOR_REGION)
                ANCHOR_OFFSET = data.get("ANCHOR_OFFSET", ANCHOR_OFFSET)
                ANCHOR_THRESHOLD = data.get("ANCHOR_THRESHOLD", ANCHOR_THRESHOLD)
                ANCHOR_TEMPLATE_DIR = data.get("ANCHOR_TEMPLATE_DIR", ANCHOR_TEMPLATE_DIR)
                ALIGNMENT_POLL_INTERVAL_MS = data.get("ALIGNMENT_POLL_INTERVAL_MS", ALIGNMENT_POLL_INTERVAL_MS)
                CONTINUOUS_CAPTURE_INTERVAL = data.get("CONTINUOUS_CAPTURE_INTERVAL", CONTINUOUS_CAPTURE_INTERVAL)
                OLLAMA_KEEP_ALIVE = data.get("OLLAMA_KEEP_ALIVE", OLLAMA_KEEP_ALIVE)
                OLLAMA_NUM_PREDICT = int(data.get("OLLAMA_NUM_PREDICT", OLLAMA_NUM_PREDICT))
                OLLAMA_NUM_PREDICT = max(16, min(64, OLLAMA_NUM_PREDICT))
                OLLAMA_REQUEST_TIMEOUT = max(8.0, min(120.0, float(data.get("OLLAMA_REQUEST_TIMEOUT", OLLAMA_REQUEST_TIMEOUT))))
                OLLAMA_FALLBACK_ENABLED = bool(data.get("OLLAMA_FALLBACK_ENABLED", OLLAMA_FALLBACK_ENABLED))
                FAST_OCR_ENABLED = bool(data.get("FAST_OCR_ENABLED", FAST_OCR_ENABLED))
                FAST_OCR_MIN_CONFIDENCE = max(0.0, min(1.0, float(data.get("FAST_OCR_MIN_CONFIDENCE", FAST_OCR_MIN_CONFIDENCE))))
                INFO_OVERLAY_OFFSET = data.get("INFO_OVERLAY_OFFSET", INFO_OVERLAY_OFFSET)
                MATCH_BOX_ENABLED = bool(data.get("MATCH_BOX_ENABLED", MATCH_BOX_ENABLED))
                MATCH_BOX_OFFSET = data.get("MATCH_BOX_OFFSET", MATCH_BOX_OFFSET)
                MATCH_BOX_SCALE = max(0.6, min(2.0, float(data.get("MATCH_BOX_SCALE", MATCH_BOX_SCALE))))
                MATCH_BOX_OPACITY = max(0.2, min(1.0, float(data.get("MATCH_BOX_OPACITY", MATCH_BOX_OPACITY))))
                configured_host = sanitize_ollama_host(data.get("OLLAMA_HOST", CONFIGURED_OLLAMA_HOST))
                if configured_host != CONFIGURED_OLLAMA_HOST:
                    CONFIGURED_OLLAMA_HOST = configured_host
                    if CONFIGURED_OLLAMA_HOST:
                        os.environ["OLLAMA_HOST"] = CONFIGURED_OLLAMA_HOST
                    reset_ollama_client()
        except (json.JSONDecodeError, OSError) as e:
            logger.warning(f"Config file invalid or empty, resetting: {e}")
            save_config()
    else:
        save_config()

    ensure_anchor_directory(ANCHOR_TEMPLATE_DIR)
    last_alignment_info["enabled"] = AUTO_ALIGN_ENABLED



def save_config():
    global CAP_REGION, label_color, AUTO_ALIGN_ENABLED, ANCHOR_REGION, ANCHOR_OFFSET, ANCHOR_THRESHOLD, ANCHOR_TEMPLATE_DIR
    global ALIGNMENT_POLL_INTERVAL_MS, CONTINUOUS_CAPTURE_INTERVAL, INFO_OVERLAY_OFFSET, MATCH_BOX_ENABLED, MATCH_BOX_OFFSET, MATCH_BOX_SCALE, MATCH_BOX_OPACITY, CONFIGURED_OLLAMA_HOST
    global OLLAMA_KEEP_ALIVE, OLLAMA_NUM_PREDICT, OLLAMA_REQUEST_TIMEOUT, OLLAMA_FALLBACK_ENABLED, FAST_OCR_ENABLED, FAST_OCR_MIN_CONFIDENCE
    data = {
        "CAP_REGION": CAP_REGION,
        "label_color": label_color,
        "AUTO_ALIGN_ENABLED": AUTO_ALIGN_ENABLED,
        "ANCHOR_REGION": ANCHOR_REGION,
        "ANCHOR_OFFSET": ANCHOR_OFFSET,
        "ANCHOR_THRESHOLD": ANCHOR_THRESHOLD,
        "ANCHOR_TEMPLATE_DIR": ANCHOR_TEMPLATE_DIR,
        "ALIGNMENT_POLL_INTERVAL_MS": ALIGNMENT_POLL_INTERVAL_MS,
        "CONTINUOUS_CAPTURE_INTERVAL": CONTINUOUS_CAPTURE_INTERVAL,
        "OLLAMA_KEEP_ALIVE": OLLAMA_KEEP_ALIVE,
        "OLLAMA_NUM_PREDICT": OLLAMA_NUM_PREDICT,
        "OLLAMA_REQUEST_TIMEOUT": OLLAMA_REQUEST_TIMEOUT,
        "OLLAMA_FALLBACK_ENABLED": OLLAMA_FALLBACK_ENABLED,
        "FAST_OCR_ENABLED": FAST_OCR_ENABLED,
        "FAST_OCR_MIN_CONFIDENCE": FAST_OCR_MIN_CONFIDENCE,
        "INFO_OVERLAY_OFFSET": INFO_OVERLAY_OFFSET,
        "MATCH_BOX_ENABLED": MATCH_BOX_ENABLED,
        "MATCH_BOX_OFFSET": MATCH_BOX_OFFSET,
        "MATCH_BOX_SCALE": MATCH_BOX_SCALE,
        "MATCH_BOX_OPACITY": MATCH_BOX_OPACITY,
        "OLLAMA_HOST": CONFIGURED_OLLAMA_HOST,
    }
    with open(CONFIG_FILE, "w") as f:
        json.dump(data, f, indent=4)
    logger.info("Config saved.")


# ---------- Load Ore Reference Data ----------

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller bundle."""
    if hasattr(sys, "_MEIPASS"):  # running inside PyInstaller bundle
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

USER_ROCK_DATA = {"materials": [], "sightings": []}
STARMINERS_ORE_DATA = {"ores": []}


def default_user_rock_data() -> Dict[str, Dict]:
    return {"materials": [], "sightings": [], "STANTON": {}, "PYRO": {}}


def load_starminers_ore_data() -> Dict:
    path = resource_path(STARMINERS_ORE_DATA_FILE)
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError) as exc:
        logger.warning(f"Unable to load Star Miners ore data from {path}: {exc}")
        return {"ores": []}
    if not isinstance(data, dict):
        return {"ores": []}
    ores = data.get("ores", [])
    if not isinstance(ores, list):
        data["ores"] = []
    return data


def normalize_ore_record(ore: Dict) -> Dict:
    tier_colors = {"S": "#E88AFF", "A": "#63E64C", "B": "#E6E14C", "C": "#E69E4C"}
    name = str(ore.get("name", "")).strip()
    types = ore.get("ty", ore.get("types", []))
    locations = ore.get("l", ore.get("locations", []))
    rs_value = ore.get("rs")
    try:
        rs_value = int(rs_value) if rs_value not in (None, "") else None
    except (TypeError, ValueError):
        rs_value = None
    return {
        "name": name,
        "color": ore.get("c") or tier_colors.get(str(ore.get("t", "")).upper(), "#888"),
        "price": int(ore.get("p") or 0),
        "price_label": f"{int(ore.get('p') or 0):,}" if ore.get("p") else "Unknown",
        "tier": str(ore.get("t") or "C").upper(),
        "types": types if isinstance(types, list) else [],
        "volatile": bool(ore.get("v")),
        "rs": rs_value,
        "craft": ore.get("craft") or "-",
        "note": ore.get("n") or "",
        "locations": locations if isinstance(locations, list) else [],
        "source": ore.get("source") or "starminers",
    }


def get_ore_reference_rows() -> List[Dict]:
    rows_by_name: Dict[str, Dict] = {}
    for ore in STARMINERS_ORE_DATA.get("ores", []):
        if not isinstance(ore, dict):
            continue
        normalized = normalize_ore_record(ore)
        if normalized["name"]:
            rows_by_name[normalized["name"].upper()] = normalized

    for ore in USER_ROCK_DATA.get("materials", []):
        if not isinstance(ore, dict):
            continue
        normalized = normalize_ore_record({**ore, "source": "local"})
        if normalized["name"]:
            rows_by_name[normalized["name"].upper()] = normalized

    return sorted(rows_by_name.values(), key=lambda item: item["name"].upper())


def _parse_numeric_code(numeric_code: Optional[str]) -> Optional[int]:
    if not numeric_code:
        return None
    try:
        match = re.search(r"\d+", numeric_code.replace(",", ""))
    except AttributeError:
        return None
    if not match:
        return None
    try:
        return int(match.group(0))
    except ValueError:
        return None


def find_ore_matches(*, text: Optional[str] = None, numeric_code: Optional[str] = None) -> List[Dict]:
    rows = get_ore_reference_rows()
    haystack = (text or "").upper()
    value = _parse_numeric_code(numeric_code)
    for ore in rows:
        name = ore["name"].upper()
        if name and re.search(rf"\b{re.escape(name)}\b", haystack):
            matched = dict(ore)
            matched["match_type"] = "name"
            matched["scan_rs"] = value
            base_rs = ore.get("rs")
            if isinstance(base_rs, int) and base_rs > 0 and value and value % base_rs == 0:
                matched["units"] = min(50, max(1, value // base_rs))
            else:
                matched["units"] = 1
            matched["additional_units"] = max(0, matched["units"] - 1)
            matched["display_units"] = matched["units"]
            return [matched]

    if value is not None:
        exact_matches = []
        for ore in rows:
            base_rs = ore.get("rs")
            if not isinstance(base_rs, int) or base_rs <= 0:
                continue
            if value % base_rs == 0:
                units = value // base_rs
                if 1 <= units <= 50:
                    matched = dict(ore)
                    matched["match_type"] = "multiple"
                    matched["scan_rs"] = value
                    matched["units"] = units
                    matched["additional_units"] = max(0, units - 1)
                    matched["display_units"] = matched["units"]
                    exact_matches.append(matched)
        if exact_matches:
            exact_matches.sort(key=lambda ore: (ore["units"], ore["name"]))
            return exact_matches

        with_rs = [ore for ore in rows if isinstance(ore.get("rs"), int)]
        if with_rs:
            nearest = min(with_rs, key=lambda ore: abs(int(ore["rs"]) - value))
            if abs(int(nearest["rs"]) - value) <= 8:
                matched = dict(nearest)
                matched["match_type"] = "near"
                matched["scan_rs"] = value
                matched["units"] = 1
                matched["additional_units"] = 0
                matched["display_units"] = 1
                return [matched]
    return []


def find_near_ore_matches(numeric_code: Optional[str], *, max_abs_diff: int = 250, max_pct_diff: float = 0.015) -> List[Dict]:
    """Return close RS total candidates without treating them as exact ore matches."""

    value = _parse_numeric_code(numeric_code)
    if value is None:
        return []

    candidates: List[Dict] = []
    for ore in get_ore_reference_rows():
        base_rs = ore.get("rs")
        if not isinstance(base_rs, int) or base_rs <= 0:
            continue
        for units in range(1, 51):
            expected = base_rs * units
            diff = abs(expected - value)
            allowed = max_abs_diff if expected < 10000 else max(max_abs_diff, int(round(expected * max_pct_diff)))
            if diff <= allowed:
                matched = dict(ore)
                matched["match_type"] = "near_multiple"
                matched["scan_rs"] = value
                matched["expected_rs"] = expected
                matched["difference"] = diff
                matched["units"] = units
                matched["additional_units"] = max(0, units - 1)
                matched["display_units"] = units
                candidates.append(matched)

    candidates.sort(key=lambda ore: (ore["difference"], ore["units"], ore["name"]))
    return candidates[:6]


def find_ore_reference(*, text: Optional[str] = None, numeric_code: Optional[str] = None) -> Optional[Dict]:
    matches = find_ore_matches(text=text, numeric_code=numeric_code)
    if matches:
        return matches[0]
    return None


def quality_advice(value: Optional[int]) -> Optional[Dict]:
    if value is None:
        return None
    value = max(0, min(1000, int(value)))
    if value < 500:
        return {
            "value": value,
            "grade": "Below standard",
            "outcome": "Worse than store-bought",
            "bonus": "Negative stat penalty",
            "recommendation": "Sell raw at refinery",
            "rarity": "Common",
            "color": "#ff6b5a",
        }
    if value < 700:
        return {
            "value": value,
            "grade": "Standard grade",
            "outcome": "Equal to store-bought",
            "bonus": "Minimal (+1-3%)",
            "recommendation": "Refine for bulk crafting",
            "rarity": "Fairly common",
            "color": "#E6E14C",
        }
    if value < 900:
        return {
            "value": value,
            "grade": "High grade",
            "outcome": "Noticeably better",
            "bonus": "Good (+4-8% per stat)",
            "recommendation": "Keep and refine for crafting",
            "rarity": "Uncommon",
            "color": "#63E64C",
        }
    return {
        "value": value,
        "grade": "Premium / Elite",
        "outcome": "Maximum possible stats",
        "bonus": "Excellent (+8-12%+)",
        "recommendation": "Hold for peak crafting value",
        "rarity": "Very rare",
        "color": "#E88AFF",
    }


def extract_quality_from_text(raw_text: str) -> Optional[int]:
    if not raw_text:
        return None
    explicit = re.search(r"(?:quality|qual|q)\D{0,8}(\d{1,4})", raw_text, re.IGNORECASE)
    if explicit:
        value = int(explicit.group(1))
        if 0 <= value <= 1000:
            return value
    return None


def ensure_user_rock_data_file() -> None:
    if os.path.exists(USER_ROCK_DATA_FILE):
        return
    with open(USER_ROCK_DATA_FILE, "w") as f:
        json.dump(default_user_rock_data(), f, indent=4)


def load_user_rock_data() -> Dict:
    ensure_user_rock_data_file()
    try:
        with open(USER_ROCK_DATA_FILE, "r") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError) as exc:
        logger.warning(f"User rock data invalid or unreadable, resetting: {exc}")
        data = default_user_rock_data()
        save_user_rock_data(data)

    if not isinstance(data, dict):
        data = default_user_rock_data()
    data.setdefault("materials", [])
    if not isinstance(data["materials"], list):
        data["materials"] = []
    data.setdefault("sightings", [])
    if not isinstance(data["sightings"], list):
        data["sightings"] = []
    for region in ("STANTON", "PYRO"):
        data.setdefault(region, {})
    return data


def save_user_rock_data(data: Dict) -> None:
    tmp_file = f"{USER_ROCK_DATA_FILE}.tmp"
    with open(tmp_file, "w") as f:
        json.dump(data, f, indent=4, sort_keys=True)
    os.replace(tmp_file, USER_ROCK_DATA_FILE)


def reload_data_sources() -> None:
    global USER_ROCK_DATA, STARMINERS_ORE_DATA
    with DATA_LOCK:
        STARMINERS_ORE_DATA = load_starminers_ore_data()
        USER_ROCK_DATA = load_user_rock_data()


def save_user_material_entry(payload: Dict) -> Dict:
    system_name = str(payload.get("system") or payload.get("region") or "Stanton").strip().title()
    if system_name.upper() not in {"STANTON", "PYRO", "NYX", "OTHER"}:
        raise ValueError("System must be Stanton, Pyro, Nyx, or Other.")

    ore_name = str(payload.get("ore_name", "")).strip()
    if not ore_name:
        raise ValueError("Ore name is required.")

    location = str(payload.get("location", "")).strip()
    if not location:
        raise ValueError("Location is required.")

    location_type = str(payload.get("location_type") or "asteroid_belt").strip().lower()
    quality_raw = payload.get("quality")
    quality_seen = None
    if quality_raw not in (None, ""):
        try:
            quality_seen = max(0, min(1000, int(float(quality_raw))))
        except (TypeError, ValueError):
            raise ValueError("Quality must be a number from 0 to 1000.")

    ore_ref = find_ore_reference(text=ore_name)
    if ore_ref:
        ore_name = ore_ref["name"]

    rs_raw = payload.get("rs")
    rs_value = None
    if rs_raw not in (None, ""):
        try:
            rs_value = max(1, int(float(str(rs_raw).replace(",", ""))))
        except (TypeError, ValueError):
            raise ValueError("RS must be a positive number.")

    sighting = {
        "ore_name": ore_name,
        "system": system_name,
        "location": location,
        "location_type": location_type,
        "quality": quality_seen,
        "quality_advice": quality_advice(quality_seen),
        "notes": str(payload.get("notes", "")).strip(),
        "source": "local",
        "game_version": "4.8",
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
    }

    with DATA_LOCK:
        data = load_user_rock_data()
        if rs_value is not None:
            materials = data.setdefault("materials", [])
            material = {
                "name": ore_name,
                "c": "#8fd7ff",
                "p": 0,
                "t": "C",
                "ty": ["ship"],
                "v": False,
                "rs": rs_value,
                "craft": "User discovered",
                "n": "Local user material entry",
                "l": [location],
                "source": "local",
            }
            existing_index = next(
                (idx for idx, item in enumerate(materials) if str(item.get("name", "")).strip().upper() == ore_name.upper()),
                None,
            )
            if existing_index is None:
                materials.append(material)
            else:
                previous_locations = materials[existing_index].get("l", [])
                if isinstance(previous_locations, list) and location not in previous_locations:
                    material["l"] = previous_locations + [location]
                materials[existing_index].update(material)
        data.setdefault("sightings", []).append(sighting)
        save_user_rock_data(data)
        reload_data_sources()

    return sighting


reload_data_sources()

# ---------- OCR with Ollama ----------
def prepare_ocr_image(pil_img: Image.Image) -> Image.Image:
    """Normalize scan crops to a small, high-contrast image for fast VLM OCR."""

    img = pil_img.convert("RGB")
    width, height = img.size
    if width <= 0 or height <= 0:
        return img

    target_height = 56
    scale = target_height / float(height)
    if scale < 1.0 or height < 44:
        new_width = max(1, int(round(width * scale)))
        img = img.resize((new_width, target_height), Image.Resampling.BICUBIC)

    width, height = img.size
    if width > 280:
        new_height = max(1, int(round(height * (280 / float(width)))))
        img = img.resize((280, new_height), Image.Resampling.BICUBIC)

    img = ImageEnhance.Contrast(img).enhance(1.8)
    img = ImageEnhance.Sharpness(img).enhance(1.4)
    return img


def ocr_with_fast_engine(pil_img: Image.Image) -> Tuple[str, float]:
    """Read the scan crop with a lightweight OCR engine before falling back to Ollama."""

    engine = get_fast_ocr_engine()
    if engine is None:
        return "", 0.0

    img = pil_img.convert("RGB")
    width, height = img.size
    if height < 80:
        scale = max(2, int(round(96 / max(1, height))))
        img = img.resize((width * scale, height * scale), Image.Resampling.BICUBIC)

    cv_img = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    try:
        result, _ = engine(cv_img)
    except Exception as exc:
        logger.warning("Fast OCR error: %s", exc)
        return "", 0.0

    best_text = ""
    best_conf = 0.0
    for item in result or []:
        if not isinstance(item, (list, tuple)) or len(item) < 3:
            continue
        text = str(item[1] or "").strip()
        try:
            conf = float(item[2])
        except (TypeError, ValueError):
            conf = 0.0
        if conf > best_conf and re.search(r"\d", text):
            best_text = text
            best_conf = conf

    if best_conf >= float(FAST_OCR_MIN_CONFIDENCE):
        return best_text, best_conf
    return "", best_conf


def ocr_with_ollama(pil_img: Image.Image, model=OLLAMA_MODEL) -> str:
    global last_ocr_error
    last_ocr_error = None
    buf = io.BytesIO()
    prepared_img = prepare_ocr_image(pil_img)
    prepared_img.save(buf, format="PNG")
    img_bytes = buf.getvalue()
    client = get_ollama_client()
    try:
        response = client.chat(
            model=model,
            messages=[{
                "role": "user",
                "content": (
                    "OCR only. Return visible ore name, RS digits, and quality digits if present. "
                    "No explanation."
                ),
                "images": [img_bytes],
            }],
            options={
                "temperature": 0,
                "num_predict": int(OLLAMA_NUM_PREDICT),
                "num_ctx": 1024,
            },
            keep_alive=OLLAMA_KEEP_ALIVE,
        )
        message = response.get("message", {})
        if isinstance(message, dict):
            content = (message.get("content") or "").strip()
            thinking = (message.get("thinking") or "").strip()
        else:
            content = (getattr(message, "content", "") or "").strip()
            thinking = (getattr(message, "thinking", "") or "").strip()
        if content:
            return content
        if thinking:
            logger.warning("Ollama returned empty content with thinking text; ignoring thinking output.")
        return ""
    except Exception as e:
        last_ocr_error = str(e)
        logger.error(f"Ollama OCR error: {e}")
        return ""


def extract_code_from_text(raw_text: str):
    if not raw_text:
        return None, None
    rs_match = re.search(r"\bRS\s*[:#-]?\s*(\d[\d,\.]{1,10})\b", raw_text, re.IGNORECASE)
    if rs_match:
        raw_rs = rs_match.group(0)
        return rs_match.group(1).replace(",", "").replace(".", ""), raw_rs
    matches = CODE_RE.findall(raw_text)
    if not matches:
        return None, raw_text
    raw = matches[0].upper()
    if any(ch.isdigit() for ch in raw):
        m = re.match(r"([A-Za-z]?-?)([\d,\.]+)", raw)
        if m:
            prefix, digits = m.groups()
            digits = digits.replace(",", "").replace(".", "")
            candidate = prefix + digits
        else:
            candidate = raw.replace(",", "").replace(".", "")
    else:
        candidate = raw
    return candidate, raw


# ---------- Capture / Overlay ----------
continuous_mode = False
show_border = True
show_keybind_overlay = False
border_canvas = None

capture_overlay_root = None
capture_overlay_canvas = None
capture_rect_id = None
capture_overlay_animation_job = None
capture_overlay_last_layout = {
    "overlay_width": None,
    "overlay_height": None,
    "left": None,
    "top": None,
    "cap_w": None,
    "cap_h": None,
}

info_overlay_root = None
info_overlay_canvas = None
info_text_id = None
info_keybind_text_id = None
info_overlay_geometry = {"screen_width": None, "screen_height": None, "width": 0, "height": 0}
overlay_text = ""
last_overlay_time = 0

match_box_root = None
match_box_canvas = None
match_box_geometry = {"screen_width": None, "screen_height": None, "width": 0, "height": 0}

anchor_overlay_root = None
anchor_overlay_canvas = None
anchor_rect_id = None
anchor_overlay_visible = True


def perform_auto_alignment() -> bool:
    """Attempt to adjust the capture region based on anchor template matches."""
    global CAP_REGION, last_alignment_info

    last_alignment_info["enabled"] = AUTO_ALIGN_ENABLED

    if not AUTO_ALIGN_ENABLED:
        return False

    if anchor_tracker is None:
        logger.debug("Anchor tracker not initialised; skipping auto alignment.")
        return False

    anchor_tracker.set_threshold(float(ANCHOR_THRESHOLD))
    detection = anchor_tracker.locate_anchor(ANCHOR_REGION)

    if not detection:
        last_alignment_info.update({
            "matched": False,
            "template": None,
            "score": 0.0,
            "match_left": None,
            "match_top": None,
            "capture_left": None,
            "capture_top": None,
        })
        return False

    template_w = detection.get("template_width", float(CAP_REGION["width"]))
    template_h = detection.get("template_height", float(CAP_REGION["height"]))
    base_left = detection["match_left"] + (template_w / 2.0) - (CAP_REGION["width"] / 2.0)
    base_top = detection["match_top"] + (template_h / 2.0) - (CAP_REGION["height"] / 2.0)

    new_left = int(round(base_left + ANCHOR_OFFSET.get("x", 0)))
    new_top = int(round(base_top + ANCHOR_OFFSET.get("y", 0)))

    CAP_REGION["left"] = max(0, new_left)
    CAP_REGION["top"] = max(0, new_top)

    last_alignment_info.update({
        "matched": True,
        "template": detection["template"],
        "score": float(detection["score"]),
        "match_left": detection["match_left"],
        "match_top": detection["match_top"],
        "capture_left": CAP_REGION["left"],
        "capture_top": CAP_REGION["top"],
    })

    sync_capture_sliders()

    if capture_overlay_root:
        try:
            capture_overlay_root.after(0, update_capture_overlay_region)
        except (RuntimeError, tk.TclError):
            update_capture_overlay_region()

    logger.debug(
        "Auto alignment applied using %s (score %.3f) => CAP_REGION left/top updated to (%d, %d)",
        detection["template"],
        detection["score"],
        CAP_REGION["left"],
        CAP_REGION["top"],
    )
    return True


def toggle_border():
    """Toggle visibility of the debug red border."""
    global show_border, border_canvas
    show_border = not show_border
    if border_canvas:
        border_canvas.itemconfig("border", state="normal" if show_border else "hidden")
    refresh_toggle_buttons()


def update_keybind_overlay() -> None:
    if not info_overlay_canvas or not info_keybind_text_id:
        return
    text = "7  Scan   Ctrl+7  Loop   8  Border   Ctrl+8  Keys" if show_keybind_overlay else ""
    info_overlay_canvas.itemconfig(info_keybind_text_id, text=text)


def toggle_keybind_overlay():
    """Toggle the keybind reference on the in-game result overlay."""
    global show_keybind_overlay
    show_keybind_overlay = not show_keybind_overlay
    update_keybind_overlay()
    refresh_toggle_buttons()


def format_ore_display_name(info: Optional[Dict]) -> str:
    if not info:
        return ""
    name = str(info.get("name", "")).strip()
    if not name:
        return ""

    display_units = info.get("display_units", info.get("units", 1))
    try:
        count = int(display_units)
    except (TypeError, ValueError):
        count = 1

    if count > 1:
        return f"{name} x{count}"
    return name


def update_overlay_label(info, *, code: Optional[str] = None, raw_text: Optional[str] = None) -> None:
    """Update the floating label with the latest scan result."""
    global overlay_text, info_overlay_canvas, info_text_id, last_overlay_time

    message = format_ore_display_name(info)

    overlay_text = message
    if message:
        last_overlay_time = time.time()
    else:
        last_overlay_time = 0
    if info_overlay_canvas and info_text_id:
        info_overlay_canvas.itemconfig(info_text_id, text=overlay_text, fill=label_color)
    draw_match_box()


def compute_match_box_geometry(screen_width: int, screen_height: int) -> Tuple[int, int, int, int]:
    scale = max(0.6, min(2.0, float(MATCH_BOX_SCALE)))
    width = int(round(420 * scale))
    height = int(round(580 * scale))
    base_left = max(0, (screen_width - width) // 2)
    base_top = max(0, (screen_height - height) // 2)
    left = min(max(0, base_left + int(MATCH_BOX_OFFSET.get("x", 0))), max(0, screen_width - width))
    top = min(max(0, base_top + int(MATCH_BOX_OFFSET.get("y", 0))), max(0, screen_height - height))
    return width, height, left, top


def _hex_to_rgb(color: str, fallback: Tuple[int, int, int] = (94, 229, 255)) -> Tuple[int, int, int]:
    try:
        color = str(color).strip().lstrip("#")
        if len(color) == 3:
            color = "".join(ch * 2 for ch in color)
        return int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
    except Exception:
        return fallback


def _blend_hex(color: str, amount: float = 0.35) -> str:
    r, g, b = _hex_to_rgb(color)
    r = min(255, int(r + (255 - r) * amount))
    g = min(255, int(g + (255 - g) * amount))
    b = min(255, int(b + (255 - b) * amount))
    return f"#{r:02x}{g:02x}{b:02x}"


def _draw_round_rect(canvas: tk.Canvas, x1: int, y1: int, x2: int, y2: int, radius: int, **kwargs) -> None:
    radius = max(1, min(radius, (x2 - x1) // 2, (y2 - y1) // 2))
    canvas.create_rectangle(x1 + radius, y1, x2 - radius, y2, **kwargs)
    canvas.create_rectangle(x1, y1 + radius, x2, y2 - radius, **kwargs)
    canvas.create_oval(x1, y1, x1 + radius * 2, y1 + radius * 2, **kwargs)
    canvas.create_oval(x2 - radius * 2, y1, x2, y1 + radius * 2, **kwargs)
    canvas.create_oval(x1, y2 - radius * 2, x1 + radius * 2, y2, **kwargs)
    canvas.create_oval(x2 - radius * 2, y2 - radius * 2, x2, y2, **kwargs)


def _draw_pill(canvas: tk.Canvas, x: int, y: int, text: str, fill: str, scale: float) -> int:
    font_size = max(7, int(round(10 * scale)))
    width = max(int(round(58 * scale)), int(round((len(text) * 7 + 18) * scale)))
    height = int(round(22 * scale))
    _draw_round_rect(canvas, x, y, x + width, y + height, height // 2, fill=fill, outline=fill)
    canvas.create_text(
        x + width // 2,
        y + height // 2,
        text=text.upper(),
        fill="#06131f",
        font=("Arial", font_size, "bold"),
    )
    return width


def _match_box_price(ore: Dict) -> str:
    price = ore.get("price")
    try:
        price_value = int(price)
    except (TypeError, ValueError):
        price_value = 0
    if price_value <= 0:
        return "Unknown"
    types = ore.get("types") or []
    unit = " /unit" if isinstance(types, list) and types and all(t in {"roc", "fps"} for t in types) else " /SCU"
    return f"{price_value:,}{unit}"


def draw_match_box() -> None:
    if not match_box_canvas:
        return
    width = int(match_box_geometry.get("width") or 260)
    height = int(match_box_geometry.get("height") or 112)
    scale = max(0.6, min(2.0, float(MATCH_BOX_SCALE)))
    match_box_canvas.delete("all")
    if not MATCH_BOX_ENABLED:
        return
    result = last_match_result if isinstance(last_match_result, dict) else None
    ore = result.get("ore") if result else None
    if not isinstance(ore, dict):
        return

    pad = int(round(20 * scale))
    accent = ore.get("color") or label_color
    muted = "#8bc8e5"
    text = "#def7ff"
    table_bg = "#0b1a2d"
    border = "#58d4ff"

    _draw_round_rect(match_box_canvas, 2, 2, width - 2, height - 2, int(round(8 * scale)), fill="#10263f", outline=border, width=1)
    match_box_canvas.create_polygon(
        int(width * 0.68),
        0,
        width,
        0,
        int(width * 0.28),
        height,
        0,
        height,
        fill="#1a4963",
        stipple="gray75",
    )
    match_box_canvas.create_polygon(
        0,
        0,
        int(width * 0.48),
        0,
        int(width * 0.12),
        height,
        0,
        height,
        fill="#315e76",
        stipple="gray50",
    )

    title = format_ore_display_name(ore).upper()
    match_box_canvas.create_text(
        pad,
        pad,
        text=title,
        fill=label_color,
        font=("Arial", max(16, int(round(28 * scale))), "bold"),
        anchor="nw",
        width=width - (pad * 2),
    )
    scan_rs = ore.get("scan_rs") or result.get("code")
    match_box_canvas.create_text(
        pad,
        pad + int(round(52 * scale)),
        text=f"SCANNED RS: {scan_rs}" if scan_rs else "SCANNED RS: UNKNOWN",
        fill=muted,
        font=("Arial", max(9, int(round(13 * scale))), "bold"),
        anchor="nw",
    )

    pill_y = pad + int(round(84 * scale))
    pill_x = pad
    for label, color in (("Highest", "#E88AFF"), ("High", "#63E64C"), ("Medium", "#E6E14C"), ("Low", "#E69E4C")):
        pill_x += _draw_pill(match_box_canvas, pill_x, pill_y, label, color, scale) + int(round(6 * scale))

    table_x = pad
    table_y = pill_y + int(round(36 * scale))
    table_w = width - (pad * 2)
    header_h = int(round(34 * scale))
    row_h = int(round(39 * scale))
    rows = [
        ("Ore", format_ore_display_name(ore)),
        ("Scanned RS", str(scan_rs or "Unknown")),
        ("Base RS", str(ore.get("rs") or "Unknown")),
        ("Amount", f"{int(ore.get('display_units') or ore.get('units') or 1)} detected"),
        ("Refined Value", _match_box_price(ore)),
        ("Tier / Type", f"{ore.get('tier') or 'C'} / {', '.join(ore.get('types') or ['ship']).upper()}"),
        ("Crafting Use", str(ore.get("craft") or "-")),
    ]
    table_h = header_h + (row_h * len(rows))
    _draw_round_rect(match_box_canvas, table_x, table_y, table_x + table_w, table_y + table_h, int(round(8 * scale)), fill=table_bg, outline="#205374", width=1)
    match_box_canvas.create_rectangle(table_x, table_y, table_x + table_w, table_y + header_h, fill="#0d4770", outline="#205374")
    match_box_canvas.create_text(
        table_x + table_w // 2,
        table_y + header_h // 2,
        text="STAR MINERS DEPOT ORE REFERENCE",
        fill=text,
        font=("Arial", max(8, int(round(11 * scale))), "bold"),
    )
    split_x = table_x + int(round(table_w * 0.34))
    y = table_y + header_h
    for label, value in rows:
        match_box_canvas.create_line(table_x, y, table_x + table_w, y, fill="#205374")
        match_box_canvas.create_line(split_x, y, split_x, y + row_h, fill="#1c3d5a")
        match_box_canvas.create_text(
            table_x + int(round(14 * scale)),
            y + row_h // 2,
            text=label.upper(),
            fill=muted,
            font=("Arial", max(7, int(round(10 * scale)))),
            anchor="w",
            width=split_x - table_x - int(round(20 * scale)),
        )
        match_box_canvas.create_text(
            split_x + (table_x + table_w - split_x) // 2,
            y + row_h // 2,
            text=value,
            fill=text if label != "Ore" else _blend_hex(accent, 0.45),
            font=("Arial", max(8, int(round(11 * scale))), "bold"),
            justify="center",
            width=table_x + table_w - split_x - int(round(16 * scale)),
        )
        y += row_h

    chip_y = table_y + table_h + int(round(12 * scale))
    chip_x = pad
    max_chip_y = height - int(round(44 * scale))
    for location in (ore.get("locations") or [])[:10]:
        chip_w = max(int(round(70 * scale)), int(round((len(str(location)) * 6 + 20) * scale)))
        chip_h = int(round(22 * scale))
        if chip_x + chip_w > width - pad:
            chip_x = pad
            chip_y += chip_h + int(round(6 * scale))
        if chip_y + chip_h > max_chip_y:
            break
        _draw_round_rect(match_box_canvas, chip_x, chip_y, chip_x + chip_w, chip_y + chip_h, chip_h // 2, fill="#5eddf5", outline="#5eddf5")
        match_box_canvas.create_text(
            chip_x + chip_w // 2,
            chip_y + chip_h // 2,
            text=str(location).upper(),
            fill="#06131f",
            font=("Arial", max(7, int(round(9 * scale))), "bold"),
        )
        chip_x += chip_w + int(round(6 * scale))

    note = str(ore.get("note") or "").strip()
    if note:
        match_box_canvas.create_text(
            pad,
            height - int(round(30 * scale)),
            text=note,
            fill="#c0d0dc",
            font=("Arial", max(7, int(round(10 * scale))), "bold"),
            anchor="w",
            width=width - (pad * 2),
        )


def show_match_box_overlay(screen_width: Optional[int] = None, screen_height: Optional[int] = None) -> None:
    global match_box_root, match_box_canvas
    if screen_width is None or screen_height is None:
        screen_width = 1920
        screen_height = 1080
        try:
            if info_overlay_root and info_overlay_root.winfo_exists():
                screen_width = info_overlay_root.winfo_screenwidth()
                screen_height = info_overlay_root.winfo_screenheight()
        except tk.TclError:
            pass
    width, height, left, top = compute_match_box_geometry(screen_width, screen_height)
    if match_box_root and match_box_root.winfo_exists():
        match_box_root.geometry(f"{width}x{height}+{left}+{top}")
        if match_box_canvas:
            match_box_canvas.config(width=width, height=height)
    else:
        match_box_root = create_overlay_window(width, height, left, top)
        match_box_canvas = tk.Canvas(match_box_root, width=width, height=height, bg="black", highlightthickness=0)
        match_box_canvas.pack()
        enforce_topmost(match_box_root)
    try:
        match_box_root.attributes("-alpha", max(0.2, min(1.0, float(MATCH_BOX_OPACITY))))
    except tk.TclError:
        pass
    match_box_geometry.update({"screen_width": screen_width, "screen_height": screen_height, "width": width, "height": height})
    draw_match_box()


def hide_match_box_overlay() -> None:
    global match_box_root, match_box_canvas
    if match_box_root and match_box_root.winfo_exists():
        try:
            match_box_root.destroy()
        except tk.TclError:
            pass
    match_box_root = None
    match_box_canvas = None


def reposition_match_box_overlay() -> None:
    if MATCH_BOX_ENABLED:
        show_match_box_overlay()
    else:
        hide_match_box_overlay()


def toggle_match_box_overlay() -> None:
    global MATCH_BOX_ENABLED
    MATCH_BOX_ENABLED = not MATCH_BOX_ENABLED
    reposition_match_box_overlay()
    refresh_toggle_buttons()


def compute_info_overlay_geometry(screen_width: int, screen_height: int) -> Tuple[int, int, int, int]:
    overlay_width = max(400, min(800, screen_width - 40))
    overlay_height = 120
    base_left = max(0, (screen_width - overlay_width) // 2)
    base_top = max(0, int(screen_height * 0.35) - overlay_height // 2)

    offset_x = int(INFO_OVERLAY_OFFSET.get("x", 0))
    offset_y = int(INFO_OVERLAY_OFFSET.get("y", 0))

    max_left = max(0, screen_width - overlay_width)
    max_top = max(0, screen_height - overlay_height)

    left = min(max(0, base_left + offset_x), max_left)
    top = min(max(0, base_top + offset_y), max_top)
    return overlay_width, overlay_height, left, top


def reposition_info_overlay() -> None:
    global info_overlay_canvas, info_text_id, info_keybind_text_id
    if not info_overlay_root or not info_overlay_canvas or not info_text_id:
        return
    if not info_overlay_root.winfo_exists():
        return

    try:
        screen_width = info_overlay_root.winfo_screenwidth()
        screen_height = info_overlay_root.winfo_screenheight()
    except tk.TclError:
        geo = info_overlay_geometry
        screen_width = geo.get("screen_width") or 1920
        screen_height = geo.get("screen_height") or 1080

    overlay_width, overlay_height, left, top = compute_info_overlay_geometry(screen_width, screen_height)

    info_overlay_root.geometry(f"{overlay_width}x{overlay_height}+{left}+{top}")
    info_overlay_canvas.config(width=overlay_width, height=overlay_height)
    info_overlay_canvas.coords(info_text_id, overlay_width // 2, overlay_height // 2 - 6)
    info_overlay_canvas.itemconfig(info_text_id, width=overlay_width - 60)
    if info_keybind_text_id:
        info_overlay_canvas.coords(info_keybind_text_id, overlay_width // 2, overlay_height - 18)
        info_overlay_canvas.itemconfig(info_keybind_text_id, width=overlay_width - 60)

    info_overlay_geometry.update(
        {
            "screen_width": screen_width,
            "screen_height": screen_height,
            "width": overlay_width,
            "height": overlay_height,
        }
    )


def start_label_timeout(window: Optional[tk.Toplevel]) -> None:
    """Background loop to clear overlay label if no update for 10s."""
    global info_overlay_canvas, info_text_id, last_overlay_time

    if info_overlay_canvas and info_text_id:
        if last_overlay_time and (time.time() - last_overlay_time > 10):
            info_overlay_canvas.itemconfig(info_text_id, text="")
            last_overlay_time = 0

    if window and window.winfo_exists():
        window.after(500, lambda: start_label_timeout(window))



def enforce_topmost(window: tk.Toplevel, interval_ms: int = 1500) -> None:
    """Continuously lift the overlay window so it stays above focused apps."""
    if window is None:
        return
    if not window.winfo_exists():
        return
    try:
        window.attributes("-topmost", True)
        window.lift()
    except tk.TclError:
        return
    window.after(interval_ms, lambda: enforce_topmost(window, interval_ms))


def create_overlay_window(width: int, height: int, left: int, top: int) -> tk.Toplevel:
    """Create a transparent always-on-top overlay window."""
    window = tk.Toplevel()
    window.attributes("-transparentcolor", "black")
    window.attributes("-topmost", True)
    window.overrideredirect(True)
    window.configure(bg="black")
    window.geometry(f"{width}x{height}+{left}+{top}")
    enforce_topmost(window)
    return window



# ---------- GUI + Overlay ----------
def choose_label_color():
    global label_color, info_overlay_canvas, info_text_id
    color = colorchooser.askcolor(title="Choose Label Color")[1]
    if color:
        label_color = color
        if info_overlay_canvas and info_text_id:
            info_overlay_canvas.itemconfig(info_text_id, fill=label_color)


def _compute_capture_overlay_layout() -> Dict[str, int]:
    cap_w = int(CAP_REGION['width'])
    cap_h = int(CAP_REGION['height'])
    padding_x, padding_y = 100, 40

    overlay_width = cap_w + padding_x
    overlay_height = cap_h + padding_y
    left = int(CAP_REGION['left']) - (padding_x // 2)
    top = int(CAP_REGION['top']) - padding_y

    return {
        "overlay_width": overlay_width,
        "overlay_height": overlay_height,
        "left": left,
        "top": top,
        "padding_x": padding_x,
        "padding_y": padding_y,
        "cap_w": cap_w,
        "cap_h": cap_h,
    }


def _apply_capture_overlay_layout(*, force: bool = False) -> None:
    global capture_overlay_last_layout

    if not capture_overlay_canvas or not capture_rect_id or not capture_overlay_root:
        return

    layout = _compute_capture_overlay_layout()

    padding_x = layout["padding_x"]
    padding_y = layout["padding_y"]
    cap_w = layout["cap_w"]
    cap_h = layout["cap_h"]
    overlay_width = layout["overlay_width"]
    overlay_height = layout["overlay_height"]
    left = layout["left"]
    top = layout["top"]

    last = capture_overlay_last_layout
    size_changed = (
        force
        or last["overlay_width"] != overlay_width
        or last["overlay_height"] != overlay_height
    )
    pos_changed = force or last["left"] != left or last["top"] != top
    rect_changed = force or last["cap_w"] != cap_w or last["cap_h"] != cap_h

    if size_changed:
        capture_overlay_canvas.config(width=overlay_width, height=overlay_height)

    if rect_changed:
        capture_overlay_canvas.coords(
            capture_rect_id,
            padding_x // 2,
            padding_y,
            padding_x // 2 + cap_w,
            padding_y + cap_h,
        )

    if size_changed or pos_changed:
        capture_overlay_root.geometry(f"{overlay_width}x{overlay_height}+{left}+{top}")
        try:
            capture_overlay_root.lift()
        except tk.TclError:
            pass

    last.update(
        {
            "overlay_width": overlay_width,
            "overlay_height": overlay_height,
            "left": left,
            "top": top,
            "cap_w": cap_w,
            "cap_h": cap_h,
        }
    )


def _animate_capture_overlay() -> None:
    global capture_overlay_animation_job

    if (
        not capture_overlay_root
        or not capture_overlay_canvas
        or not capture_rect_id
        or not capture_overlay_root.winfo_exists()
    ):
        capture_overlay_animation_job = None
        return

    try:
        _apply_capture_overlay_layout()
    except tk.TclError:
        capture_overlay_animation_job = None
        return

    try:
        capture_overlay_animation_job = capture_overlay_root.after(33, _animate_capture_overlay)
    except tk.TclError:
        capture_overlay_animation_job = None


def start_capture_overlay_animation(*, force: bool = False) -> None:
    global capture_overlay_animation_job

    if not capture_overlay_root or not capture_overlay_canvas or not capture_rect_id:
        return

    _apply_capture_overlay_layout(force=force)

    if capture_overlay_animation_job is None:
        try:
            capture_overlay_animation_job = capture_overlay_root.after(33, _animate_capture_overlay)
        except tk.TclError:
            capture_overlay_animation_job = None


def stop_capture_overlay_animation() -> None:
    global capture_overlay_animation_job, capture_overlay_last_layout

    if capture_overlay_animation_job is not None and capture_overlay_root:
        try:
            capture_overlay_root.after_cancel(capture_overlay_animation_job)
        except (tk.TclError, ValueError):
            pass
    capture_overlay_animation_job = None

    capture_overlay_last_layout.update(
        {
            "overlay_width": None,
            "overlay_height": None,
            "left": None,
            "top": None,
            "cap_w": None,
            "cap_h": None,
        }
    )


def show_capture_overlay():
    global border_canvas, capture_overlay_canvas, capture_overlay_root, capture_rect_id

    if capture_overlay_root and capture_overlay_root.winfo_exists():
        try:
            stop_capture_overlay_animation()
            capture_overlay_root.destroy()
        except tk.TclError:
            pass
        capture_overlay_canvas = None
        capture_rect_id = None
        border_canvas = None

    cap_w, cap_h = int(CAP_REGION['width']), int(CAP_REGION['height'])
    padding_x, padding_y = 100, 40

    overlay_width = cap_w + padding_x
    overlay_height = cap_h + padding_y
    left = int(CAP_REGION['left']) - (padding_x // 2)
    top = int(CAP_REGION['top']) - padding_y

    capture_overlay_root = create_overlay_window(overlay_width, overlay_height, left, top)

    capture_overlay_canvas = tk.Canvas(
        capture_overlay_root,
        width=overlay_width,
        height=overlay_height,
        bg="black",
        highlightthickness=0,
    )
    capture_overlay_canvas.pack()
    border_canvas = capture_overlay_canvas

    capture_rect_id = capture_overlay_canvas.create_rectangle(
        padding_x // 2,
        padding_y,
        padding_x // 2 + cap_w,
        padding_y + cap_h,
        outline="red",
        width=3,
        tags=("border",),
    )

    start_capture_overlay_animation(force=True)


def show_info_overlay(screen_width: int, screen_height: int) -> None:
    """Display a floating status overlay near the top-center of the screen."""
    global info_overlay_root, info_overlay_canvas, info_text_id, info_keybind_text_id

    if info_overlay_root and info_overlay_root.winfo_exists():
        try:
            info_overlay_root.destroy()
        except tk.TclError:
            pass
        info_overlay_canvas = None
        info_text_id = None
        info_keybind_text_id = None

    overlay_width, overlay_height, left, top = compute_info_overlay_geometry(screen_width, screen_height)

    info_overlay_root = create_overlay_window(overlay_width, overlay_height, left, top)

    info_overlay_canvas = tk.Canvas(
        info_overlay_root,
        width=overlay_width,
        height=overlay_height,
        bg="black",
        highlightthickness=0,
    )
    info_overlay_canvas.pack()

    info_text_id = info_overlay_canvas.create_text(
        overlay_width // 2,
        overlay_height // 2 - 6,
        text=overlay_text,
        fill=label_color,
        font=("Arial", 18, "bold"),
        width=overlay_width - 60,
        justify="center",
    )
    info_keybind_text_id = info_overlay_canvas.create_text(
        overlay_width // 2,
        overlay_height - 18,
        text="",
        fill="#9ee8ff",
        font=("Arial", 10, "bold"),
        width=overlay_width - 60,
        justify="center",
    )
    update_keybind_overlay()

    info_overlay_geometry.update(
        {
            "screen_width": screen_width,
            "screen_height": screen_height,
            "width": overlay_width,
            "height": overlay_height,
        }
    )

    start_label_timeout(info_overlay_root)


def update_capture_overlay_region():
    start_capture_overlay_animation(force=True)


def show_anchor_overlay():
    global anchor_overlay_root, anchor_overlay_canvas, anchor_rect_id, anchor_overlay_visible

    if not anchor_overlay_visible:
        return

    if anchor_overlay_root and anchor_overlay_root.winfo_exists():
        try:
            anchor_overlay_root.destroy()
        except tk.TclError:
            pass
        anchor_overlay_canvas = None
        anchor_rect_id = None

    pad = 40
    width = int(ANCHOR_REGION['width']) + pad
    height = int(ANCHOR_REGION['height']) + pad
    left = int(ANCHOR_REGION['left']) - (pad // 2)
    top = int(ANCHOR_REGION['top']) - (pad // 2)

    anchor_overlay_root = create_overlay_window(width, height, left, top)

    anchor_overlay_canvas = tk.Canvas(
        anchor_overlay_root,
        width=width,
        height=height,
        bg="black",
        highlightthickness=0,
    )
    anchor_overlay_canvas.pack()

    anchor_rect_id = anchor_overlay_canvas.create_rectangle(
        pad // 2,
        pad // 2,
        pad // 2 + int(ANCHOR_REGION['width']),
        pad // 2 + int(ANCHOR_REGION['height']),
        outline="#00d4ff",
        width=2,
    )

    anchor_overlay_canvas.create_text(
        width // 2,
        5,
        text="ANCHOR REGION",
        fill="#00d4ff",
        font=("Arial", 12, "bold"),
        anchor="n",
    )


def update_anchor_overlay_region():
    global anchor_overlay_root, anchor_overlay_canvas, anchor_rect_id
    if (
        not anchor_overlay_visible
        or not anchor_overlay_root
        or not anchor_overlay_canvas
        or not anchor_rect_id
    ):
        return

    pad = 40
    width = int(ANCHOR_REGION['width']) + pad
    height = int(ANCHOR_REGION['height']) + pad
    left = int(ANCHOR_REGION['left']) - (pad // 2)
    top = int(ANCHOR_REGION['top']) - (pad // 2)

    anchor_overlay_canvas.config(width=width, height=height)

    anchor_overlay_canvas.coords(
        anchor_rect_id,
        pad // 2,
        pad // 2,
        pad // 2 + int(ANCHOR_REGION['width']),
        pad // 2 + int(ANCHOR_REGION['height']),
    )
    anchor_overlay_root.geometry(f"{width}x{height}+{left}+{top}")
    try:
        anchor_overlay_root.lift()
    except tk.TclError:
        pass


def hide_anchor_overlay():
    global anchor_overlay_root, anchor_overlay_canvas, anchor_rect_id

    if anchor_overlay_root and anchor_overlay_root.winfo_exists():
        try:
            anchor_overlay_root.destroy()
        except tk.TclError:
            pass

    anchor_overlay_root = None
    anchor_overlay_canvas = None
    anchor_rect_id = None


def show_overlay(screen_width: Optional[int] = None, screen_height: Optional[int] = None) -> None:
    show_anchor_overlay()
    show_capture_overlay()

    if screen_width is None or screen_height is None:
        try:
            if capture_overlay_root and capture_overlay_root.winfo_exists():
                screen_width = capture_overlay_root.winfo_screenwidth()
                screen_height = capture_overlay_root.winfo_screenheight()
        except tk.TclError:
            screen_width = screen_width or 1920
            screen_height = screen_height or 1080

    if screen_width is not None and screen_height is not None:
        show_info_overlay(screen_width, screen_height)
        if MATCH_BOX_ENABLED:
            show_match_box_overlay(screen_width, screen_height)


def update_overlay_region():
    update_anchor_overlay_region()
    update_capture_overlay_region()
    reposition_match_box_overlay()


def launch_gui():
    def update_region_from_sliders(*args):
        if GUI_CONTROL_STATE["syncing"]["capture"]:
            return
        CAP_REGION["left"] = int(slider_left.get())
        CAP_REGION["top"] = int(slider_top.get())
        CAP_REGION["width"] = int(slider_width.get())
        CAP_REGION["height"] = int(slider_height.get())
        status_var.set(f"CAP_REGION updated: {CAP_REGION}")
        update_capture_overlay_region()

    def update_anchor_region_from_sliders(*args):
        if GUI_CONTROL_STATE["syncing"]["anchor"]:
            return
        ANCHOR_REGION["left"] = int(anchor_left.get())
        ANCHOR_REGION["top"] = int(anchor_top.get())
        ANCHOR_REGION["width"] = int(anchor_width.get())
        ANCHOR_REGION["height"] = int(anchor_height.get())
        set_anchor_status(f"Anchor region updated: {ANCHOR_REGION}", hold=2.0)
        if AUTO_ALIGN_ENABLED:
            perform_auto_alignment()
        update_anchor_overlay_region()

    def update_anchor_offset_from_sliders(*args):
        if GUI_CONTROL_STATE["syncing"]["anchor"]:
            return
        ANCHOR_OFFSET["x"] = int(anchor_offset_x.get())
        ANCHOR_OFFSET["y"] = int(anchor_offset_y.get())
        set_anchor_status(f"Anchor offset updated: {ANCHOR_OFFSET}", hold=2.0)
        if AUTO_ALIGN_ENABLED:
            perform_auto_alignment()

    def update_info_overlay_from_sliders(*args):
        if GUI_CONTROL_STATE["syncing"].get("overlay"):
            return
        INFO_OVERLAY_OFFSET["x"] = int(info_offset_x.get())
        INFO_OVERLAY_OFFSET["y"] = int(info_offset_y.get())
        status_var.set(
            f"Display offset updated: x={INFO_OVERLAY_OFFSET['x']}, y={INFO_OVERLAY_OFFSET['y']}"
        )
        reposition_info_overlay()

    def update_match_box_from_sliders(*args):
        global MATCH_BOX_SCALE, MATCH_BOX_OPACITY
        MATCH_BOX_OFFSET["x"] = int(match_box_offset_x.get())
        MATCH_BOX_OFFSET["y"] = int(match_box_offset_y.get())
        MATCH_BOX_SCALE = max(0.6, min(2.0, float(match_box_size.get())))
        MATCH_BOX_OPACITY = max(0.2, min(1.0, float(match_box_opacity.get())))
        status_var.set(
            f"Match box updated: x={MATCH_BOX_OFFSET['x']}, y={MATCH_BOX_OFFSET['y']}, size={MATCH_BOX_SCALE:.2f}, opacity={MATCH_BOX_OPACITY:.2f}"
        )
        reposition_match_box_overlay()

    def toggle_auto_align():
        global AUTO_ALIGN_ENABLED
        AUTO_ALIGN_ENABLED = auto_align_var.get()
        last_alignment_info["enabled"] = AUTO_ALIGN_ENABLED
        if AUTO_ALIGN_ENABLED:
            set_anchor_status("Head sway compensation enabled.")
            perform_auto_alignment()
        else:
            set_anchor_status("Head sway compensation disabled.")

    def reload_anchor_templates():
        global anchor_tracker
        ensure_anchor_directory(ANCHOR_TEMPLATE_DIR)
        if anchor_tracker is None:
            anchor_tracker = AnchorRegionTracker(ANCHOR_TEMPLATE_DIR, ANCHOR_THRESHOLD)
        count = anchor_tracker.set_directory(ANCHOR_TEMPLATE_DIR)
        set_anchor_status(f"Loaded {count} anchor template(s) from {ANCHOR_TEMPLATE_DIR}.")

    def manual_realign():
        success = perform_auto_alignment()
        if success:
            set_anchor_status(
                f"Anchor locked using {last_alignment_info['template']} (score {last_alignment_info['score']:.2f}).",
                hold=2.5,
            )
            status_var.set(f"Auto alignment adjusted CAP_REGION: {CAP_REGION}")
        else:
            set_anchor_status("Anchor match not found. Adjust search region or add templates.")

    def open_anchor_directory():
        path = os.path.abspath(ANCHOR_TEMPLATE_DIR)
        ensure_anchor_directory(path)
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform.startswith("darwin"):
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as exc:
            set_anchor_status(f"Unable to open template folder: {exc}", hold=3.0)
        else:
            set_anchor_status(f"Opened template folder: {path}")

    def update_threshold(*_args):
        global ANCHOR_THRESHOLD
        try:
            value = float(threshold_var.get())
        except (tk.TclError, ValueError):
            return
        value = max(0.1, min(0.99, value))
        ANCHOR_THRESHOLD = value
        if anchor_tracker is not None:
            anchor_tracker.set_threshold(ANCHOR_THRESHOLD)
        set_anchor_status(f"Anchor detection threshold set to {ANCHOR_THRESHOLD:.2f}")

    def toggle_anchor_overlay_visibility():
        global anchor_overlay_visible
        anchor_overlay_visible = anchor_overlay_var.get()
        if anchor_overlay_visible:
            show_anchor_overlay()
            set_anchor_status("Anchor overlay shown.")
        else:
            hide_anchor_overlay()
            set_anchor_status("Anchor overlay hidden.")

    def update_alignment_interval(*_args):
        global ALIGNMENT_POLL_INTERVAL_MS
        try:
            value = int(alignment_interval_var.get())
        except (tk.TclError, ValueError):
            return
        value = max(100, min(5000, value))
        ALIGNMENT_POLL_INTERVAL_MS = value
        set_anchor_status(f"Alignment interval set to {ALIGNMENT_POLL_INTERVAL_MS} ms", hold=2.0)

    def update_capture_interval(*_args):
        global CONTINUOUS_CAPTURE_INTERVAL
        try:
            value = float(capture_interval_var.get())
        except (tk.TclError, ValueError):
            return
        value = max(0.2, min(30.0, value))
        CONTINUOUS_CAPTURE_INTERVAL = value
        status_var.set(f"Continuous capture interval set to {CONTINUOUS_CAPTURE_INTERVAL:.1f}s")

    def alignment_poll():
        now = time.time()
        message: Optional[str] = None

        if AUTO_ALIGN_ENABLED:
            if anchor_tracker is None or not getattr(anchor_tracker, "templates", None):
                message = "Add anchor templates to enable head sway compensation."
                last_alignment_info.update(
                    {
                        "matched": False,
                        "template": None,
                        "score": 0.0,
                        "match_left": None,
                        "match_top": None,
                        "capture_left": None,
                        "capture_top": None,
                    }
                )
            else:
                match_found = perform_auto_alignment()
                info = last_alignment_info
                if info.get("matched"):
                    message = (
                        f"Anchor locked using {info['template']} (score {info['score']:.2f})."
                    )
                    capture_msg = f"Auto alignment adjusted CAP_REGION: {CAP_REGION}"
                    if status_var.get() != capture_msg:
                        status_var.set(capture_msg)
                elif not match_found:
                    message = "Anchor match not found. Adjust search region or add templates."
        else:
            message = "Head sway compensation disabled."
            last_alignment_info.update(
                {
                    "matched": False,
                    "template": None,
                    "score": 0.0,
                    "match_left": None,
                    "match_top": None,
                    "capture_left": None,
                    "capture_top": None,
                }
            )

        if message and now >= anchor_status_hold["until"]:
            if message != alignment_status_cache.get("message") or anchor_status_var.get() != message:
                anchor_status_var.set(message)
                alignment_status_cache["message"] = message

        try:
            interval = max(100, int(ALIGNMENT_POLL_INTERVAL_MS))
            root.after(interval, alignment_poll)
        except tk.TclError:
            pass

    def on_close():
        global capture_overlay_root, capture_overlay_canvas, capture_rect_id
        global anchor_overlay_root, anchor_overlay_canvas, anchor_rect_id
        global info_overlay_root, info_overlay_canvas, info_text_id, info_keybind_text_id
        global match_box_root, match_box_canvas

        stop_capture_overlay_animation()

        save_config()
        try:
            for window in (capture_overlay_root, anchor_overlay_root, info_overlay_root, match_box_root):
                if window and window.winfo_exists():
                    window.destroy()
        except Exception:
            pass

        capture_overlay_root = None
        capture_overlay_canvas = None
        capture_rect_id = None
        anchor_overlay_root = None
        anchor_overlay_canvas = None
        anchor_rect_id = None
        info_overlay_root = None
        info_overlay_canvas = None
        info_text_id = None
        info_keybind_text_id = None
        match_box_root = None
        match_box_canvas = None

        root.destroy()

    root = tk.Tk()
    root.title("Star Citizen Scanner Control")
    root.protocol("WM_DELETE_WINDOW", on_close)

    colors = apply_glass_theme(root)

    status_var = tk.StringVar(value="Ready.")
    anchor_status_var = tk.StringVar(value="Head sway compensation ready.")
    ollama_host_var = tk.StringVar(value=CONFIGURED_OLLAMA_HOST)
    ollama_active_host_var = tk.StringVar()

    def refresh_active_host_label() -> None:
        ollama_active_host_var.set(f"Active host: {get_ollama_host()}")

    def apply_ollama_host_from_ui() -> None:
        sanitized = set_configured_ollama_host(ollama_host_var.get())
        refresh_active_host_label()
        active_host = get_ollama_host()
        if sanitized:
            ollama_host_var.set(sanitized)
            message = f"Remote Ollama host set to {active_host}."
        else:
            message = f"Ollama host cleared. Using {active_host}."
        status_var.set(message)
        logger.info(message)

    def use_local_ollama_host() -> None:
        ollama_host_var.set("")
        set_configured_ollama_host("")
        refresh_active_host_label()
        active_host = get_ollama_host()
        message = f"Ollama host cleared. Using {active_host}."
        status_var.set(message)
        logger.info(message)

    def open_mobile_overlay() -> None:
        url = f"http://{get_local_ip()}:5000"
        try:
            webbrowser.open_new_tab(url)
        except Exception as exc:
            status_var.set(f"Unable to open browser: {exc}")
            logger.warning("Failed to open mobile overlay URL %s: %s", url, exc)
        else:
            status_var.set(f"Opening overlay in browser: {url}")
            logger.info("Opened mobile overlay URL: %s", url)

    alignment_status_cache = {"message": None}
    anchor_status_hold = {"until": 0.0}

    def set_anchor_status(message: str, hold: float = 1.5) -> None:
        anchor_status_var.set(message)
        anchor_status_hold["until"] = time.time() + hold
        alignment_status_cache["message"] = None

    refresh_active_host_label()

    # Scrollable container so the full control panel is accessible on smaller displays.
    container = ttk.Frame(root, style="Glass.Main.TFrame")
    container.pack(fill="both", expand=True)

    canvas = tk.Canvas(
        container,
        background=colors["background"],
        highlightthickness=0,
        borderwidth=0,
    )
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    main = ttk.Frame(canvas, style="Glass.Main.TFrame", padding=20)
    window_id = canvas.create_window((15, 15), window=main, anchor="nw")

    def _sync_scroll_region(_event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfigure(window_id, width=canvas.winfo_width())

    def _on_mousewheel(event):
        # Windows/MacOS delta uses event.delta; Linux uses Button-4/5 events below.
        step = int(-1 * (event.delta / 120))
        canvas.yview_scroll(step, "units")

    def _on_mousewheel_linux(event, direction: int):
        canvas.yview_scroll(direction, "units")

    main.bind("<Configure>", _sync_scroll_region)
    canvas.bind("<Configure>", _sync_scroll_region)
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    canvas.bind_all("<Button-4>", lambda e: _on_mousewheel_linux(e, -1))
    canvas.bind_all("<Button-5>", lambda e: _on_mousewheel_linux(e, 1))

    frm_region = ttk.LabelFrame(main, text="Capture Region", style="Glass.TLabelframe")
    frm_region.pack(fill="x", padx=5, pady=8)

    slider_left = create_glass_scale(
        frm_region,
        text="Left",
        minimum=0,
        maximum=3000,
        initial=CAP_REGION["left"],
        command=update_region_from_sliders,
    )

    slider_top = create_glass_scale(
        frm_region,
        text="Top",
        minimum=0,
        maximum=2000,
        initial=CAP_REGION["top"],
        command=update_region_from_sliders,
    )

    slider_width = create_glass_scale(
        frm_region,
        text="Width",
        minimum=50,
        maximum=1000,
        initial=CAP_REGION["width"],
        command=update_region_from_sliders,
    )

    slider_height = create_glass_scale(
        frm_region,
        text="Height",
        minimum=20,
        maximum=500,
        initial=CAP_REGION["height"],
        command=update_region_from_sliders,
        padding=(0, 0),
    )

    register_capture_sliders(slider_left, slider_top, slider_width, slider_height)
    sync_capture_sliders()

    frm_anchor = ttk.LabelFrame(main, text="Head Sway Compensation", style="Glass.TLabelframe")
    frm_anchor.pack(fill="x", padx=5, pady=8)

    auto_align_var = tk.BooleanVar(value=AUTO_ALIGN_ENABLED)
    chk_auto_align = ttk.Checkbutton(
        frm_anchor,
        text="Enable auto alignment",
        variable=auto_align_var,
        command=toggle_auto_align,
        style="Glass.TCheckbutton",
    )
    chk_auto_align.pack(anchor="w", padx=5, pady=(5, 0))

    anchor_overlay_var = tk.BooleanVar(value=anchor_overlay_visible)
    chk_anchor_overlay = ttk.Checkbutton(
        frm_anchor,
        text="Show anchor overlay",
        variable=anchor_overlay_var,
        command=toggle_anchor_overlay_visibility,
        style="Glass.TCheckbutton",
    )
    chk_anchor_overlay.pack(anchor="w", padx=5, pady=(0, 5))

    interval_row = ttk.Frame(frm_anchor, style="Glass.Section.TFrame")
    interval_row.pack(fill="x", padx=5, pady=(0, 5))
    ttk.Label(interval_row, text="Alignment interval (ms)", style="Glass.Small.TLabel").pack(side="left")
    alignment_interval_var = tk.IntVar(value=int(ALIGNMENT_POLL_INTERVAL_MS))
    alignment_interval_spin = tk.Spinbox(
        interval_row,
        from_=100,
        to=5000,
        increment=50,
        textvariable=alignment_interval_var,
        width=6,
        command=update_alignment_interval,
    )
    alignment_interval_spin.pack(side="left", padx=5)
    style_spinbox(alignment_interval_spin, colors)
    alignment_interval_var.trace_add("write", update_alignment_interval)

    threshold_row = ttk.Frame(frm_anchor, style="Glass.Section.TFrame")
    threshold_row.pack(fill="x", padx=5, pady=5)
    ttk.Label(threshold_row, text="Detection threshold", style="Glass.Small.TLabel").pack(side="left")
    threshold_var = tk.DoubleVar(value=ANCHOR_THRESHOLD)
    threshold_spin = tk.Spinbox(threshold_row, from_=0.10, to=0.99, increment=0.01,
                                 textvariable=threshold_var, width=6, command=update_threshold)
    threshold_spin.pack(side="left", padx=5)
    style_spinbox(threshold_spin, colors)
    threshold_var.trace_add("write", update_threshold)

    anchor_left = create_glass_scale(
        frm_anchor,
        text="Anchor Left",
        minimum=0,
        maximum=3840,
        initial=ANCHOR_REGION["left"],
        command=update_anchor_region_from_sliders,
    )

    anchor_top = create_glass_scale(
        frm_anchor,
        text="Anchor Top",
        minimum=0,
        maximum=2160,
        initial=ANCHOR_REGION["top"],
        command=update_anchor_region_from_sliders,
    )

    anchor_width = create_glass_scale(
        frm_anchor,
        text="Anchor Width",
        minimum=50,
        maximum=1200,
        initial=ANCHOR_REGION["width"],
        command=update_anchor_region_from_sliders,
    )

    anchor_height = create_glass_scale(
        frm_anchor,
        text="Anchor Height",
        minimum=50,
        maximum=800,
        initial=ANCHOR_REGION["height"],
        command=update_anchor_region_from_sliders,
    )

    anchor_offset_x = create_glass_scale(
        frm_anchor,
        text="Offset X",
        minimum=-300,
        maximum=600,
        initial=ANCHOR_OFFSET["x"],
        command=update_anchor_offset_from_sliders,
    )

    anchor_offset_y = create_glass_scale(
        frm_anchor,
        text="Offset Y",
        minimum=-300,
        maximum=600,
        initial=ANCHOR_OFFSET["y"],
        command=update_anchor_offset_from_sliders,
        padding=(0, 0),
    )

    register_anchor_sliders(
        anchor_left,
        anchor_top,
        anchor_width,
        anchor_height,
        anchor_offset_x,
        anchor_offset_y,
    )
    sync_anchor_sliders()

    frm_network = ttk.LabelFrame(main, text="Ollama Connection", style="Glass.TLabelframe")
    frm_network.pack(fill="x", padx=5, pady=8)
    ttk.Label(
        frm_network,
        text="Remote Ollama host (IPv4/hostname with optional port). Leave blank to use this PC.",
        style="Glass.Small.TLabel",
        wraplength=360,
        justify="left",
    ).pack(fill="x", padx=5, pady=(5, 2))
    host_entry = ttk.Entry(frm_network, textvariable=ollama_host_var)
    host_entry.pack(fill="x", padx=5, pady=(0, 5))

    network_button_row = ttk.Frame(frm_network, style="Glass.Section.TFrame")
    network_button_row.pack(fill="x", padx=5, pady=(0, 5))
    ttk.Button(
        network_button_row,
        text="Apply Host",
        command=apply_ollama_host_from_ui,
        style="Glass.TButton",
    ).pack(side="left", padx=5)
    ttk.Button(
        network_button_row,
        text="Use Localhost",
        command=use_local_ollama_host,
        style="Glass.TButton",
    ).pack(side="left", padx=5)
    ttk.Button(
        network_button_row,
        text="Open Mobile UI",
        command=open_mobile_overlay,
        style="Glass.TButton",
    ).pack(side="left", padx=5)
    ttk.Label(
        frm_network,
        textvariable=ollama_active_host_var,
        style="Glass.Small.TLabel",
        justify="left",
    ).pack(fill="x", padx=5, pady=(0, 5))

    anchor_btn_row = ttk.Frame(frm_anchor, style="Glass.Section.TFrame")
    anchor_btn_row.pack(fill="x", padx=5, pady=5)
    ttk.Button(anchor_btn_row, text="Reload Templates", command=reload_anchor_templates, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(anchor_btn_row, text="Realign Now", command=manual_realign, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(anchor_btn_row, text="Open Template Folder", command=open_anchor_directory, style="Glass.TButton").pack(side="left", padx=5)

    frm_display = ttk.LabelFrame(main, text="Result Display", style="Glass.TLabelframe")
    frm_display.pack(fill="x", padx=5, pady=8)

    info_offset_x = create_glass_scale(
        frm_display,
        text="Display offset X",
        minimum=-800,
        maximum=800,
        initial=int(INFO_OVERLAY_OFFSET.get("x", 0)),
        command=update_info_overlay_from_sliders,
    )

    info_offset_y = create_glass_scale(
        frm_display,
        text="Display offset Y",
        minimum=-600,
        maximum=600,
        initial=int(INFO_OVERLAY_OFFSET.get("y", 0)),
        command=update_info_overlay_from_sliders,
        padding=(0, 0),
    )

    register_overlay_sliders(info_offset_x, info_offset_y)
    sync_overlay_sliders()

    ttk.Label(frm_display, text="In-game match box", style="Glass.Small.TLabel").pack(anchor="w", padx=8, pady=(8, 0))
    match_box_offset_x = create_glass_scale(
        frm_display,
        text="Match box X",
        minimum=-900,
        maximum=900,
        initial=int(MATCH_BOX_OFFSET.get("x", 0)),
        command=update_match_box_from_sliders,
    )

    match_box_offset_y = create_glass_scale(
        frm_display,
        text="Match box Y",
        minimum=-700,
        maximum=700,
        initial=int(MATCH_BOX_OFFSET.get("y", 0)),
        command=update_match_box_from_sliders,
    )

    match_box_size = create_glass_scale(
        frm_display,
        text="Match box size",
        minimum=0.6,
        maximum=2.0,
        initial=float(MATCH_BOX_SCALE),
        command=update_match_box_from_sliders,
        resolution=0.05,
    )

    match_box_opacity = create_glass_scale(
        frm_display,
        text="Match box opacity",
        minimum=0.2,
        maximum=1.0,
        initial=float(MATCH_BOX_OPACITY),
        command=update_match_box_from_sliders,
        resolution=0.05,
        padding=(0, 0),
    )

    frm_ctrl = ttk.LabelFrame(main, text="Controls", style="Glass.TLabelframe")
    frm_ctrl.pack(fill="x", padx=5, pady=8)

    capture_interval_frame = ttk.Frame(frm_ctrl, style="Glass.Section.TFrame")
    capture_interval_frame.pack(fill="x", padx=5, pady=(5, 10))
    ttk.Label(capture_interval_frame, text="Continuous capture interval (s)", style="Glass.Small.TLabel").pack(side="left")
    capture_interval_var = tk.DoubleVar(value=float(CONTINUOUS_CAPTURE_INTERVAL))
    capture_interval_spin = tk.Spinbox(
        capture_interval_frame,
        from_=0.2,
        to=30.0,
        increment=0.1,
        textvariable=capture_interval_var,
        width=6,
        format="%.1f",
        command=update_capture_interval,
    )
    capture_interval_spin.pack(side="left", padx=5)
    style_spinbox(capture_interval_spin, colors)
    capture_interval_var.trace_add("write", update_capture_interval)

    button_row = ttk.Frame(frm_ctrl, style="Glass.Section.TFrame")
    button_row.pack(fill="x", padx=5, pady=(0, 5))

    def add_control_button(text: str, keybind: str, command: Callable[[], None], *, style: str = "Glass.TButton") -> ttk.Button:
        control = ttk.Frame(button_row, style="Glass.Section.TFrame")
        control.pack(side="left", padx=5, anchor="n")
        button = ttk.Button(control, text=text, command=command, style=style)
        button.pack()
        ttk.Label(control, text=keybind, style="Glass.Small.TLabel").pack(pady=(2, 0))
        return button

    add_control_button("Single Scan", "7", capture_once_async)
    loop_button = add_control_button("Loop: OFF", "Ctrl+7", toggle_continuous)
    add_control_button("Update Overlay", "Manual", update_overlay_region)
    add_control_button("Set Label Color", "Manual", choose_label_color)
    add_control_button("Save Config", "Manual", save_config)
    border_button = add_control_button("Border: ON", "8", toggle_border, style="Glass.Active.TButton")
    keys_button = add_control_button("Keys: OFF", "Ctrl+8", toggle_keybind_overlay)
    match_box_button = add_control_button("Match Box: OFF", "Manual", toggle_match_box_overlay)
    GUI_TOGGLE_BUTTONS["loop"] = loop_button
    GUI_TOGGLE_BUTTONS["border"] = border_button
    GUI_TOGGLE_BUTTONS["keys"] = keys_button
    GUI_TOGGLE_BUTTONS["match_box"] = match_box_button
    refresh_toggle_buttons()

    ttk.Label(main, textvariable=status_var, anchor="w", justify="left", style="Glass.Status.TLabel").pack(
        fill="x", padx=5, pady=(8, 0)
    )
    ttk.Label(main, textvariable=anchor_status_var, anchor="w", justify="left", style="Glass.Subtle.TLabel").pack(
        fill="x", padx=5, pady=(2, 5)
    )

    root.update_idletasks()
    show_overlay(root.winfo_screenwidth(), root.winfo_screenheight())
    alignment_poll()
    root.mainloop()


# ---------- Scanning Functions ----------
def capture_once():
    """Capture one scan from CAP_REGION and update overlay."""
    global last_result, last_match_result, scan_in_progress, last_scan_started_at, last_scan_skip_log_time
    if not SCAN_LOCK.acquire(blocking=False):
        now = time.time()
        if now - last_scan_skip_log_time >= SCAN_SKIP_LOG_INTERVAL:
            running_for = now - last_scan_started_at if last_scan_started_at else 0.0
            logger.info("Scan skipped because another scan is still running (%.1fs).", running_for)
            last_scan_skip_log_time = now
        return
    scan_in_progress = True
    last_scan_started_at = time.time()
    timings: Dict[str, float] = {}
    try:
        step_started = time.perf_counter()
        auto_aligned = perform_auto_alignment()
        timings["align"] = time.perf_counter() - step_started
        if AUTO_ALIGN_ENABLED:
            logger.debug("Auto alignment %s before capture.", "succeeded" if auto_aligned else "did not match")

        step_started = time.perf_counter()
        with mss.mss() as sct:
            monitor = {
                "left": CAP_REGION["left"],
                "top": CAP_REGION["top"],
                "width": CAP_REGION["width"],
                "height": CAP_REGION["height"],
            }
            img = sct.grab(monitor)
            pil_img = Image.frombytes("RGB", img.size, img.rgb)
        timings["capture"] = time.perf_counter() - step_started

        step_started = time.perf_counter()
        ocr_engine = "rapidocr"
        raw_text, ocr_confidence = ocr_with_fast_engine(pil_img)
        if not raw_text and OLLAMA_FALLBACK_ENABLED:
            ocr_engine = "ollama"
            raw_text = ocr_with_ollama(pil_img)
            ocr_confidence = 0.0
        elif not raw_text:
            ocr_engine = "rapidocr-empty"
        timings["ocr"] = time.perf_counter() - step_started

        step_started = time.perf_counter()
        code, raw = extract_code_from_text(raw_text)
        quality_value = extract_quality_from_text(raw_text)
        ore_matches = find_ore_matches(text=raw_text, numeric_code=code)
        near_ore_matches = [] if ore_matches else find_near_ore_matches(code)
        ore_info = ore_matches[0] if ore_matches else None
        quality_info = quality_advice(quality_value)
        timings["parse"] = time.perf_counter() - step_started

        last_result = {
            "code": code,
            "code_raw": raw,
            "ore": ore_info,
            "ore_matches": ore_matches,
            "near_ore_matches": near_ore_matches,
            "quality": quality_info,
            "raw_text": raw_text,
            "ocr_error": last_ocr_error,
            "ocr_engine": ocr_engine,
            "ocr_confidence": ocr_confidence,
        }
        if ore_info:
            last_match_result = dict(last_result)
        update_overlay_label(ore_info, code=code, raw_text=raw or raw_text)
        logger.info(f"Scan result: {last_result}")
        logger.info(
            "Scan timing: total=%.2fs align=%.2fs capture=%.2fs ocr=%.2fs parse=%.3fs",
            time.time() - last_scan_started_at,
            timings.get("align", 0.0),
            timings.get("capture", 0.0),
            timings.get("ocr", 0.0),
            timings.get("parse", 0.0),
        )
    finally:
        scan_in_progress = False
        last_scan_started_at = 0.0
        SCAN_LOCK.release()


def capture_once_async():
    """Start a scan without blocking the GUI event loop."""
    global last_scan_skip_log_time
    if SCAN_LOCK.locked():
        now = time.time()
        if now - last_scan_skip_log_time >= SCAN_SKIP_LOG_INTERVAL:
            running_for = now - last_scan_started_at if last_scan_started_at else 0.0
            logger.info("Scan request ignored because another scan is still running (%.1fs).", running_for)
            last_scan_skip_log_time = now
        return
    Thread(target=capture_once, daemon=True).start()


def toggle_continuous():
    """Toggle continuous scanning mode."""
    global continuous_mode, continuous_scan_thread
    continuous_mode = not continuous_mode
    logger.info(f"Continuous mode: {continuous_mode}")
    refresh_toggle_buttons()
    if continuous_mode and (continuous_scan_thread is None or not continuous_scan_thread.is_alive()):
        continuous_scan_thread = Thread(target=continuous_scan_loop, daemon=True)
        continuous_scan_thread.start()


def continuous_scan_loop():
    """Run scans repeatedly until continuous_mode is turned off."""
    while continuous_mode:
        capture_once()
        interval = max(0.1, float(CONTINUOUS_CAPTURE_INTERVAL))
        time.sleep(interval)



# ---------- Network helpers ----------
def get_local_ip() -> str:
    """Best-effort detection of the primary local network IP address."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            ip_address = sock.getsockname()[0]
            if ip_address:
                return ip_address
    except Exception as exc:
        logger.debug(f"Unable to determine local IP automatically: {exc}")
    return "127.0.0.1"


# ---------- Flask / Hotkeys ----------
template_folder = resource_path("templates")
app = Flask(__name__, template_folder=template_folder)

@app.route("/")
def index():
    return render_template("overlay.html")



@app.route("/status")
def status():
    """Return the latest scan information for the overlay UI."""

    selected_region = request.args.get("region", "STANTON").upper()
    result = last_result or {}
    ore = result.get("ore") if isinstance(result, dict) else None

    response = {
        # Legacy keys kept for compatibility with any external tools.
        "region": CAP_REGION,
        "label_color": label_color,
        "last": last_result,
        "alignment": last_alignment_info,
        # Data consumed by the overlay web page.
        "selected_region": selected_region,
        "ore": ore,
        "ore_matches": result.get("ore_matches", []) if isinstance(result, dict) else [],
        "near_ore_matches": result.get("near_ore_matches", []) if isinstance(result, dict) else [],
        "quality": result.get("quality") if isinstance(result, dict) else None,
        "code": result.get("code") if isinstance(result, dict) else None,
        "code_raw": result.get("code_raw") if isinstance(result, dict) else None,
        "confidence": float(result.get("confidence", 0.0)) if isinstance(result, dict) else 0.0,
        "raw_text": result.get("raw_text") if isinstance(result, dict) else None,
        "ocr_error": result.get("ocr_error") if isinstance(result, dict) else None,
        "ocr_engine": result.get("ocr_engine") if isinstance(result, dict) else None,
        "ocr_confidence": result.get("ocr_confidence") if isinstance(result, dict) else None,
        "ore_reference": get_ore_reference_rows(),
        "local_sightings": load_user_rock_data().get("sightings", []),
        "scan_in_progress": scan_in_progress,
        "scan_running_for": (time.time() - last_scan_started_at) if scan_in_progress and last_scan_started_at else 0.0,
    }

    return jsonify(response)


@app.route("/user-data", methods=["GET"])
def get_user_data():
    with DATA_LOCK:
        return jsonify(load_user_rock_data())


@app.route("/user-data/material", methods=["POST"])
def post_user_material():
    payload = request.get_json(silent=True) or {}
    try:
        saved = save_user_material_entry(payload)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception as exc:
        logger.exception("Unable to save user material entry.")
        return jsonify({"ok": False, "error": str(exc)}), 500
    return jsonify({"ok": True, "saved": saved, "data": load_user_rock_data()})


def hotkey_listener():
    """Set up hotkey listeners with cross-platform error handling."""
    try:
        keyboard.add_hotkey("7", capture_once_async)
        keyboard.add_hotkey("ctrl+7", toggle_continuous)
        keyboard.add_hotkey("8", toggle_border)
        keyboard.add_hotkey("ctrl+8", toggle_keybind_overlay)
        logger.info("Hotkeys registered: '7' scan, 'Ctrl+7' loop, '8' border, 'Ctrl+8' keybind overlay")
        keyboard.wait()
    except Exception as e:
        logger.warning(f"Could not set up global hotkeys: {e}")
        logger.info("Note: Linux Support is being tested.")


# ---------- Main ----------
if __name__ == "__main__":
    load_config()
    # Ensure Ollama + model before starting
    ensure_ollama_installed()
    ensure_ollama_running()
    ensure_model_installed(OLLAMA_MODEL)
    warm_fast_ocr_engine()

    anchor_tracker = AnchorRegionTracker(ANCHOR_TEMPLATE_DIR, ANCHOR_THRESHOLD)
    Thread(target=hotkey_listener, daemon=True).start()
    local_ip = get_local_ip()
    logger.info(
        "Starting overlay server: http://127.0.0.1:5000 (this device) | "
        f"http://{local_ip}:5000 (local network)"
    )
    Thread(target=lambda: app.run(host="0.0.0.0", port=5000, debug=False), daemon=True).start()
    launch_gui()
