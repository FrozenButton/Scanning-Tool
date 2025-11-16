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
from threading import Thread
from typing import Callable, Dict, List, Optional, Tuple, Union

from PIL import Image, ImageDraw, ImageTk
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


OVERLAY_PORT = 5000
LOCAL_IPV4 = "127.0.0.1"

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

CAP_REGION = {"left": 1260, "top": 310, "width": 160, "height": 30}
ANCHOR_REGION = {"left": 1100, "top": 240, "width": 320, "height": 140}
ANCHOR_OFFSET = {"x": 36, "y": 56}
ANCHOR_THRESHOLD = 0.82
AUTO_ALIGN_ENABLED = True
ANCHOR_TEMPLATE_DIR = "assets/anchor_templates"
ALIGNMENT_POLL_INTERVAL_MS = 500
CONTINUOUS_CAPTURE_INTERVAL = 2.0
INFO_OVERLAY_OFFSET = {"x": 0, "y": 0}
label_color = "yellow"
MIN_CONFIDENCE = 0.65
DEBUG_SHOW_OVERLAY = True
OLLAMA_MODEL = "qwen2.5vl:3b"   # vision model
DEFAULT_OLLAMA_HOST = "http://127.0.0.1:11434"
CONFIGURED_OLLAMA_HOST = ""

_OLLAMA_CLIENT = None
_OLLAMA_CLIENT_HOST = ""

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
        _OLLAMA_CLIENT = ollama.Client(host=host)
        _OLLAMA_CLIENT_HOST = host
    return _OLLAMA_CLIENT


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

# Regex for codes
CODE_RE = re.compile(
    r"(?:[A-Za-z]?-?\d[\d,\.]{1,10}|\d{2,10})",
    re.IGNORECASE
)

last_result = {"code": None, "code_raw": None, "info": None,
               "confidence": 0.0, "raw_text": ""}
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
    global ALIGNMENT_POLL_INTERVAL_MS, CONTINUOUS_CAPTURE_INTERVAL, INFO_OVERLAY_OFFSET, CONFIGURED_OLLAMA_HOST
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
                INFO_OVERLAY_OFFSET = data.get("INFO_OVERLAY_OFFSET", INFO_OVERLAY_OFFSET)
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
    global ALIGNMENT_POLL_INTERVAL_MS, CONTINUOUS_CAPTURE_INTERVAL, INFO_OVERLAY_OFFSET, CONFIGURED_OLLAMA_HOST
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
        "INFO_OVERLAY_OFFSET": INFO_OVERLAY_OFFSET,
        "OLLAMA_HOST": CONFIGURED_OLLAMA_HOST,
    }
    with open(CONFIG_FILE, "w") as f:
        json.dump(data, f, indent=4)
    logger.info("Config saved.")


# ---------- Load Rock Types JSON ----------

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller bundle."""
    if hasattr(sys, "_MEIPASS"):  # running inside PyInstaller bundle
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

rock_file = resource_path("RockTypes_2025-09-16.json")
with open(rock_file, "r") as f:
    ROCK_DATA = json.load(f)


# ---------- Multiplier Codes ----------
MULTIPLIER_CODES = {
    1700: {"key": "CTYPE", "display_name": "C-Type", "rarity": "common", "category": "Rock Deposits"},
    1900: {"key": "ETYPE", "display_name": "E-Type", "rarity": "common", "category": "Rock Deposits"},
    1660: {"key": "ITYPE", "display_name": "I-Type", "rarity": "common", "category": "Rock Deposits"},
    1850: {"key": "MTYPE", "display_name": "M-Type", "rarity": "common", "category": "Rock Deposits"},
    1750: {"key": "PTYPE", "display_name": "P-Type", "rarity": "common", "category": "Rock Deposits"},
    1870: {"key": "QTYPE", "display_name": "Q-Type", "rarity": "common", "category": "Rock Deposits"},
    1720: {"key": "STYPE", "display_name": "S-Type", "rarity": "common", "category": "Rock Deposits"},
    1800: {"key": "ATACAMITE", "display_name": "Atacamite", "rarity": "common", "category": "Rock Deposits"},
    1770: {"key": "FELSIC", "display_name": "Felsic", "rarity": "common", "category": "Rock Deposits"},
    1840: {"key": "GNEISS", "display_name": "Gneiss", "rarity": "common", "category": "Rock Deposits"},
    1920: {"key": "GRANITE", "display_name": "Granite", "rarity": "common", "category": "Rock Deposits"},
    1950: {"key": "IGNEOUS", "display_name": "Igneous", "rarity": "common", "category": "Rock Deposits"},
    1790: {"key": "OBSIDIAN", "display_name": "Obsidian", "rarity": "common", "category": "Rock Deposits"},
    1820: {"key": "QUARTZITE", "display_name": "Quartzite", "rarity": "common", "category": "Rock Deposits"},
    1730: {"key": "SHALE", "display_name": "Shale", "rarity": "common", "category": "Rock Deposits"},
    620: {"key": "GEMS", "display_name": "Gems", "rarity": "common", "category": "Gems"},
    2000: {"key": "SALVAGE", "display_name": "Metal Panels", "rarity": "common", "category": "Salvage"},
}

# ---------- Ore Value Tiers ----------
ORE_TIERS = {
    "HIGHEST": {"ores": ["QUANTANIUM", "STILERON", "RICCITE"], "color": "#E88AFF"},
    "HIGH": {"ores": ["TARANITE", "BEXALITE", "GOLD"], "color": "#63E64C"},
    "MEDIUM": {"ores": ["LARANITE", "BORASE", "BERYL", "AGRICIUM", "HEPHAESTANITE"], "color": "#E6E14C"},
    "LOW": {"ores": ["TUNGSTEN", "TITANIUM", "SILICON", "IRON", "QUARTZ", "CORUNDUM", "COPPER", "TIN", "ALUMINUM", "ICE"], "color": "#E69E4C"},
}
ORE_VALUE_MAP = {}
for tier, data in ORE_TIERS.items():
    for ore in data["ores"]:
        ORE_VALUE_MAP[ore.upper()] = {"tier": tier, "color": data["color"]}


# ---------- Build Deposit Tables ----------
def build_deposit_tables(rock_data):
    deposit_tables = {}
    for deposit_name, details in rock_data.items():
        ores = details.get("ores", {})
        table = []
        for ore_name, ore_data in ores.items():
            name_up = ore_name.upper()
            value_info = ORE_VALUE_MAP.get(name_up, {"tier": "OTHER", "color": "#888"})
            table.append({
                "name": ore_name.title(),
                "prob": f"{ore_data.get('prob', 0)*100:.0f}%",
                "min": f"{ore_data.get('minPct', 0)*100:.0f}%",
                "max": f"{ore_data.get('maxPct', 0)*100:.0f}%",
                "med": f"{ore_data.get('medPct', 0)*100:.0f}%",
                "tier": value_info["tier"],
                "color": value_info["color"]
            })
        tier_order = ["HIGHEST", "HIGH", "MEDIUM", "LOW", "OTHER"]
        table.sort(key=lambda x: tier_order.index(x["tier"]))
        deposit_tables[deposit_name.upper()] = table
    return deposit_tables

DEPOSIT_TABLES = {
    "STANTON": build_deposit_tables(ROCK_DATA.get("STANTON", {})),
    "PYRO": build_deposit_tables(ROCK_DATA.get("PYRO", {}))
}

# ---------- OCR with Ollama ----------
def ocr_with_ollama(pil_img: Image.Image, model=OLLAMA_MODEL) -> str:
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    img_bytes = buf.getvalue()
    client = get_ollama_client()
    try:
        response = client.chat(
            model=model,
            messages=[{
                "role": "user",
                "content": "Extract the numeric code shown in this image. Only return the code, no extra words.",
                "images": [img_bytes],
            }],
        )
        return response["message"]["content"].strip()
    except Exception as e:
        logger.error(f"Ollama OCR error: {e}")
        return ""


def extract_code_from_text(raw_text: str):
    if not raw_text:
        return None, None
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


# ---------- Deposit Lookup ----------
def lookup_deposit(code: str):
    if not code:
        return None
    try:
        m = re.search(r"(\d+)$", code)
        if not m:
            return None
        num_code = int(m.group(1))
        for base_code, info in MULTIPLIER_CODES.items():
            if num_code % base_code == 0:
                deposits = num_code // base_code
                return {
                    "name": info["display_name"],
                    "key": info["key"],
                    "rarity": info["rarity"],
                    "base_code": base_code,
                    "deposits": deposits,
                    "category": info.get("category", "Ore")
                }
    except ValueError:
        pass
    return None


# ---------- Capture / Overlay ----------
continuous_mode = False
show_border = True
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
info_overlay_geometry = {"screen_width": None, "screen_height": None, "width": 0, "height": 0}
overlay_text = ""
last_overlay_time = 0

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


def update_overlay_label(info, *, code: Optional[str] = None, raw_text: Optional[str] = None) -> None:
    """Update the floating label with the latest scan result."""
    global overlay_text, info_overlay_canvas, info_text_id, last_overlay_time

    message = ""
    if info:
        message = f"{info['name']} x{info['deposits']}" if "deposits" in info else info["name"]

    overlay_text = message
    if message:
        last_overlay_time = time.time()
    else:
        last_overlay_time = 0
    if info_overlay_canvas and info_text_id:
        info_overlay_canvas.itemconfig(info_text_id, text=overlay_text, fill=label_color)


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
    global info_overlay_canvas, info_text_id
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
    info_overlay_canvas.coords(info_text_id, overlay_width // 2, overlay_height // 2)
    info_overlay_canvas.itemconfig(info_text_id, width=overlay_width - 60)

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
    global info_overlay_root, info_overlay_canvas, info_text_id

    if info_overlay_root and info_overlay_root.winfo_exists():
        try:
            info_overlay_root.destroy()
        except tk.TclError:
            pass
        info_overlay_canvas = None
        info_text_id = None

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
        overlay_height // 2,
        text=overlay_text,
        fill=label_color,
        font=("Arial", 18, "bold"),
        width=overlay_width - 60,
        justify="center",
    )

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


def update_overlay_region():
    update_anchor_overlay_region()
    update_capture_overlay_region()


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
        global info_overlay_root, info_overlay_canvas, info_text_id

        stop_capture_overlay_animation()

        save_config()
        try:
            for window in (capture_overlay_root, anchor_overlay_root, info_overlay_root):
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

        root.destroy()

    root = tk.Tk()
    root.title("Star Citizen Scanner Control")
    root.protocol("WM_DELETE_WINDOW", on_close)

    colors = apply_glass_theme(root)

    status_var = tk.StringVar(value="Ready.")
    anchor_status_var = tk.StringVar(value="Head sway compensation ready.")
    ollama_host_var = tk.StringVar(value=CONFIGURED_OLLAMA_HOST)
    ollama_active_host_var = tk.StringVar()

    local_network_ip = LOCAL_IPV4 or get_local_ip()
    loopback_url = f"http://127.0.0.1:{OVERLAY_PORT}"
    lan_url = f"http://{local_network_ip}:{OVERLAY_PORT}"

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

    def open_web_overlay() -> None:
        try:
            opened = webbrowser.open(loopback_url)
        except Exception as exc:
            status_var.set(f"Unable to open web overlay: {exc}")
            logger.warning("Unable to open overlay browser window: %s", exc)
        else:
            if opened:
                status_var.set("Opened web overlay in browser.")
            else:
                status_var.set("Attempted to open web overlay; please check your browser.")

    alignment_status_cache = {"message": None}
    anchor_status_hold = {"until": 0.0}

    def set_anchor_status(message: str, hold: float = 1.5) -> None:
        anchor_status_var.set(message)
        anchor_status_hold["until"] = time.time() + hold
        alignment_status_cache["message"] = None

    refresh_active_host_label()

    scroll_container = ttk.Frame(root, style="Glass.Main.TFrame")
    scroll_container.pack(fill="both", expand=True, padx=15, pady=15)

    canvas = tk.Canvas(
        scroll_container,
        background=colors["background"],
        highlightthickness=0,
        borderwidth=0,
    )
    canvas.pack(side="left", fill="both", expand=True)

    v_scroll = ttk.Scrollbar(scroll_container, orient="vertical", command=canvas.yview)
    v_scroll.pack(side="right", fill="y")
    canvas.configure(yscrollcommand=v_scroll.set)

    main = ttk.Frame(canvas, style="Glass.Main.TFrame", padding=20)
    main_window = canvas.create_window((0, 0), window=main, anchor="nw")

    def _update_scrollregion(_event) -> None:
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _resize_canvas(event) -> None:
        canvas.itemconfigure(main_window, width=event.width)

    def _on_mousewheel(event) -> None:
        if event.delta:
            canvas.yview_scroll(int(-event.delta / 120), "units")
        elif getattr(event, "num", None) == 4:
            canvas.yview_scroll(-1, "units")
        elif getattr(event, "num", None) == 5:
            canvas.yview_scroll(1, "units")

    main.bind("<Configure>", _update_scrollregion)
    canvas.bind("<Configure>", _resize_canvas)
    root.bind_all("<MouseWheel>", _on_mousewheel)
    root.bind_all("<Button-4>", _on_mousewheel)
    root.bind_all("<Button-5>", _on_mousewheel)

    frm_web = ttk.LabelFrame(main, text="Web Overlay", style="Glass.TLabelframe")
    frm_web.pack(fill="x", padx=5, pady=(0, 8))

    btn_web_row = ttk.Frame(frm_web, style="Glass.Section.TFrame")
    btn_web_row.pack(fill="x", padx=5, pady=(0, 5))
    ttk.Button(
        btn_web_row,
        text="Open Web UI",
        command=open_web_overlay,
        style="Glass.TButton",
    ).pack(side="left", padx=(0, 5))

    ttk.Label(
        frm_web,
        text=f"This device: {loopback_url}",
        anchor="w",
        justify="left",
        style="Glass.TLabel",
    ).pack(fill="x", padx=5, pady=(0, 2))

    ttk.Label(
        frm_web,
        text=f"Mobile / LAN: {lan_url}",
        anchor="w",
        justify="left",
        style="Glass.TLabel",
    ).pack(fill="x", padx=5, pady=(0, 2))

    ttk.Label(
        frm_web,
        text="Share the LAN URL with mobile devices on the same network.",
        anchor="w",
        justify="left",
        style="Glass.Small.TLabel",
    ).pack(fill="x", padx=5, pady=(0, 5))

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

    ttk.Button(button_row, text="Single Scan", command=capture_once, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(button_row, text="Loop Toggle", command=toggle_continuous, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(button_row, text="Update Overlay", command=update_overlay_region, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(button_row, text="Set Label Color", command=choose_label_color, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(button_row, text="Save Config", command=save_config, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(button_row, text="Toggle Border", command=toggle_border, style="Glass.TButton").pack(side="left", padx=5)

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
    global last_result
    auto_aligned = perform_auto_alignment()
    if AUTO_ALIGN_ENABLED:
        logger.debug("Auto alignment %s before capture.", "succeeded" if auto_aligned else "did not match")
    with mss.mss() as sct:
        monitor = {
            "left": CAP_REGION["left"],
            "top": CAP_REGION["top"],
            "width": CAP_REGION["width"],
            "height": CAP_REGION["height"],
        }
        img = sct.grab(monitor)
        pil_img = Image.frombytes("RGB", img.size, img.rgb)

    raw_text = ocr_with_ollama(pil_img)
    code, raw = extract_code_from_text(raw_text)
    info = lookup_deposit(code)

    last_result = {"code": code, "code_raw": raw, "info": info, "raw_text": raw_text}
    update_overlay_label(info, code=code, raw_text=raw or raw_text)
    logger.info(f"Scan result: {last_result}")


def toggle_continuous():
    """Toggle continuous scanning mode."""
    global continuous_mode
    continuous_mode = not continuous_mode
    logger.info(f"Continuous mode: {continuous_mode}")
    if continuous_mode:
        Thread(target=continuous_scan_loop, daemon=True).start()


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
    info = result.get("info") if isinstance(result, dict) else None

    table = None
    if info:
        deposit_key = (info.get("key") or info.get("name") or "").upper()
        region_tables = DEPOSIT_TABLES.get(selected_region, {})
        table = region_tables.get(deposit_key)
        category = str(info.get("category", "")).lower()
        if not table or category not in {"rock deposits", "gems"}:
            table = None

    response = {
        # Legacy keys kept for compatibility with any external tools.
        "region": CAP_REGION,
        "label_color": label_color,
        "last": last_result,
        "alignment": last_alignment_info,
        # Data consumed by the overlay web page.
        "selected_region": selected_region,
        "info": info,
        "code": result.get("code") if isinstance(result, dict) else None,
        "code_raw": result.get("code_raw") if isinstance(result, dict) else None,
        "confidence": float(result.get("confidence", 0.0)) if isinstance(result, dict) else 0.0,
        "raw_text": result.get("raw_text") if isinstance(result, dict) else None,
        "table": table,
    }

    return jsonify(response)


def hotkey_listener():
    """Set up hotkey listeners with cross-platform error handling."""
    try:
        keyboard.add_hotkey("7", capture_once)
        keyboard.add_hotkey("ctrl+7", toggle_continuous)
        keyboard.add_hotkey("8", toggle_border)
        logger.info("Hotkeys registered: '7' for single scan, 'Ctrl+7' for continuous toggle, '8' for border toggle")
        keyboard.wait()
    except Exception as e:
        logger.warning(f"Could not set up global hotkeys: {e}")
        logger.info("Note: Linux Support is being tested.")


# ---------- Main ----------
if __name__ == "__main__":
    load_config()
    # Ensure Ollama + model before starting
    ensure_ollama_installed()
    ensure_model_installed("qwen2.5vl:3b")

    anchor_tracker = AnchorRegionTracker(ANCHOR_TEMPLATE_DIR, ANCHOR_THRESHOLD)
    Thread(target=hotkey_listener, daemon=True).start()
    local_ip = get_local_ip()
    LOCAL_IPV4 = local_ip
    logger.info(
        f"Starting overlay server: http://127.0.0.1:{OVERLAY_PORT} (this device) | "
        f"http://{local_ip}:{OVERLAY_PORT} (local network)"
    )
    Thread(
        target=lambda: app.run(host="0.0.0.0", port=OVERLAY_PORT, debug=False),
        daemon=True,
    ).start()
    launch_gui()
