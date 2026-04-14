"""Glass-themed Tkinter GUI for the scanning tool."""

import logging
import os
import subprocess
import sys
import time
import webbrowser
from typing import Callable, Dict, Optional, Tuple, Union

import tkinter as tk
from tkinter import ttk, colorchooser
from PIL import Image, ImageDraw, ImageTk

from scanning_tool.state import app_state, ScaleWidget
from scanning_tool.config import save_config, ensure_anchor_directory
from scanning_tool.overlay import (
    register_capture_sliders,
    register_anchor_sliders,
    register_overlay_sliders,
    sync_capture_sliders,
    sync_anchor_sliders,
    sync_overlay_sliders,
    show_overlay,
    show_anchor_overlay,
    hide_anchor_overlay,
    update_overlay_region,
    update_capture_overlay_region,
    update_anchor_overlay_region,
    reposition_info_overlay,
    stop_capture_overlay_animation,
    toggle_border,
    choose_label_color,
)
from scanning_tool.scanning import capture_once, toggle_continuous
from scanning_tool.anchor import AnchorRegionTracker, perform_auto_alignment
from scanning_tool.ollama_service import (
    get_ollama_host,
    get_ollama_model,
    set_configured_ollama_host,
    set_configured_ollama_model,
    ensure_model_installed,
    log_model_running_status,
)
from scanning_tool.web import get_local_ip

logger = logging.getLogger("scanning_tool")


def apply_glass_theme(root: tk.Tk) -> Dict[str, str]:
    """Apply a holographic glass-inspired theme to the Tkinter UI."""

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


def launch_gui() -> None:
    """Build and run the main Tkinter control panel."""

    def update_region_from_sliders(*args):
        if app_state.gui_control_state["syncing"]["capture"]:
            return
        app_state.cap_region["left"] = int(slider_left.get())
        app_state.cap_region["top"] = int(slider_top.get())
        app_state.cap_region["width"] = int(slider_width.get())
        app_state.cap_region["height"] = int(slider_height.get())
        status_var.set(f"CAP_REGION updated: {app_state.cap_region}")
        update_capture_overlay_region()

    def update_anchor_region_from_sliders(*args):
        if app_state.gui_control_state["syncing"]["anchor"]:
            return
        app_state.anchor_region["left"] = int(anchor_left.get())
        app_state.anchor_region["top"] = int(anchor_top.get())
        app_state.anchor_region["width"] = int(anchor_width.get())
        app_state.anchor_region["height"] = int(anchor_height.get())
        set_anchor_status(f"Anchor region updated: {app_state.anchor_region}", hold=2.0)
        if app_state.auto_align_enabled:
            perform_auto_alignment()
        update_anchor_overlay_region()

    def update_anchor_offset_from_sliders(*args):
        if app_state.gui_control_state["syncing"]["anchor"]:
            return
        app_state.anchor_offset["x"] = int(anchor_offset_x.get())
        app_state.anchor_offset["y"] = int(anchor_offset_y.get())
        set_anchor_status(f"Anchor offset updated: {app_state.anchor_offset}", hold=2.0)
        if app_state.auto_align_enabled:
            perform_auto_alignment()

    def update_info_overlay_from_sliders(*args):
        if app_state.gui_control_state["syncing"].get("overlay"):
            return
        app_state.info_overlay_offset["x"] = int(info_offset_x.get())
        app_state.info_overlay_offset["y"] = int(info_offset_y.get())
        status_var.set(
            f"Display offset updated: x={app_state.info_overlay_offset['x']}, y={app_state.info_overlay_offset['y']}"
        )
        reposition_info_overlay()

    def toggle_auto_align():
        app_state.auto_align_enabled = auto_align_var.get()
        app_state.last_alignment_info["enabled"] = app_state.auto_align_enabled
        if app_state.auto_align_enabled:
            set_anchor_status("Head sway compensation enabled.")
            perform_auto_alignment()
        else:
            set_anchor_status("Head sway compensation disabled.")

    def reload_anchor_templates():
        ensure_anchor_directory(app_state.anchor_template_dir)
        if app_state.anchor_tracker is None:
            app_state.anchor_tracker = AnchorRegionTracker(
                app_state.anchor_template_dir, app_state.anchor_threshold
            )
        count = app_state.anchor_tracker.set_directory(app_state.anchor_template_dir)
        set_anchor_status(f"Loaded {count} anchor template(s) from {app_state.anchor_template_dir}.")

    def manual_realign():
        success = perform_auto_alignment()
        if success:
            set_anchor_status(
                f"Anchor locked using {app_state.last_alignment_info['template']} "
                f"(score {app_state.last_alignment_info['score']:.2f}).",
                hold=2.5,
            )
            status_var.set(f"Auto alignment adjusted CAP_REGION: {app_state.cap_region}")
        else:
            set_anchor_status("Anchor match not found. Adjust search region or add templates.")

    def open_anchor_directory():
        path = os.path.abspath(app_state.anchor_template_dir)
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
        try:
            value = float(threshold_var.get())
        except (tk.TclError, ValueError):
            return
        value = max(0.1, min(0.99, value))
        app_state.anchor_threshold = value
        if app_state.anchor_tracker is not None:
            app_state.anchor_tracker.set_threshold(app_state.anchor_threshold)
        set_anchor_status(f"Anchor detection threshold set to {app_state.anchor_threshold:.2f}")

    def toggle_anchor_overlay_visibility():
        app_state.anchor_overlay_visible = anchor_overlay_var.get()
        if app_state.anchor_overlay_visible:
            show_anchor_overlay()
            set_anchor_status("Anchor overlay shown.")
        else:
            hide_anchor_overlay()
            set_anchor_status("Anchor overlay hidden.")

    def update_alignment_interval(*_args):
        try:
            value = int(alignment_interval_var.get())
        except (tk.TclError, ValueError):
            return
        value = max(100, min(5000, value))
        app_state.alignment_poll_interval_ms = value
        set_anchor_status(f"Alignment interval set to {app_state.alignment_poll_interval_ms} ms", hold=2.0)

    def update_capture_interval(*_args):
        try:
            value = float(capture_interval_var.get())
        except (tk.TclError, ValueError):
            return
        value = max(0.2, min(30.0, value))
        app_state.continuous_capture_interval = value
        status_var.set(f"Continuous capture interval set to {app_state.continuous_capture_interval:.1f}s")

    def alignment_poll():
        now = time.time()
        message: Optional[str] = None

        if app_state.auto_align_enabled:
            if app_state.anchor_tracker is None or not getattr(app_state.anchor_tracker, "templates", None):
                message = "Add anchor templates to enable head sway compensation."
                app_state.last_alignment_info.update({
                    "matched": False, "template": None, "score": 0.0,
                    "match_left": None, "match_top": None,
                    "capture_left": None, "capture_top": None,
                })
            else:
                match_found = perform_auto_alignment()
                info = app_state.last_alignment_info
                if info.get("matched"):
                    message = f"Anchor locked using {info['template']} (score {info['score']:.2f})."
                    capture_msg = f"Auto alignment adjusted CAP_REGION: {app_state.cap_region}"
                    if status_var.get() != capture_msg:
                        status_var.set(capture_msg)
                elif not match_found:
                    message = "Anchor match not found. Adjust search region or add templates."
        else:
            message = "Head sway compensation disabled."
            app_state.last_alignment_info.update({
                "matched": False, "template": None, "score": 0.0,
                "match_left": None, "match_top": None,
                "capture_left": None, "capture_top": None,
            })

        if message and now >= anchor_status_hold["until"]:
            if message != alignment_status_cache.get("message") or anchor_status_var.get() != message:
                anchor_status_var.set(message)
                alignment_status_cache["message"] = message

        try:
            interval = max(100, int(app_state.alignment_poll_interval_ms))
            root.after(interval, alignment_poll)
        except tk.TclError:
            pass

    def on_close():
        stop_capture_overlay_animation()
        save_config()
        try:
            for window in (app_state.capture_overlay_root, app_state.anchor_overlay_root, app_state.info_overlay_root):
                if window and window.winfo_exists():
                    window.destroy()
        except Exception:
            pass

        app_state.capture_overlay_root = None
        app_state.capture_overlay_canvas = None
        app_state.capture_rect_id = None
        app_state.anchor_overlay_root = None
        app_state.anchor_overlay_canvas = None
        app_state.anchor_rect_id = None
        app_state.info_overlay_root = None
        app_state.info_overlay_canvas = None
        app_state.info_text_id = None

        root.destroy()

    root = tk.Tk()
    root.title("Star Citizen Scanner Control")
    root.protocol("WM_DELETE_WINDOW", on_close)

    colors = apply_glass_theme(root)

    status_var = tk.StringVar(value="Ready.")
    anchor_status_var = tk.StringVar(value="Head sway compensation ready.")
    ollama_host_var = tk.StringVar(value=app_state.configured_ollama_host)
    ollama_model_var = tk.StringVar(value=app_state.ollama_model)
    ollama_active_host_var = tk.StringVar()
    ollama_active_model_var = tk.StringVar(value=f"Active model: {get_ollama_model()}")

    def refresh_active_host_label() -> None:
        ollama_active_host_var.set(f"Active host: {get_ollama_host()}")

    def refresh_active_model_label() -> None:
        ollama_active_model_var.set(f"Active model: {get_ollama_model()}")

    def apply_ollama_model_from_ui() -> None:
        model_value = ollama_model_var.get().strip()
        if not model_value:
            status_var.set("Please specify an Ollama model.")
            return
        set_configured_ollama_model(model_value)
        try:
            ensure_model_installed(model_value, exit_on_error=False)
        except Exception as exc:
            status_var.set(f"Model install failed: {exc}")
            logger.error("Failed to install model %s: %s", model_value, exc)
            return
        running = log_model_running_status(model_value)
        refresh_active_model_label()
        if running:
            status_var.set(f"Ollama model set to {model_value} and is currently running.")
        else:
            status_var.set(
                f"Ollama model set to {model_value}. It is not running yet and will start on first scan."
            )
        logger.info("Ollama model set to %s.", model_value)

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

    alignment_status_cache: Dict[str, Optional[str]] = {"message": None}
    anchor_status_hold: Dict[str, float] = {"until": 0.0}

    def set_anchor_status(message: str, hold: float = 1.5) -> None:
        anchor_status_var.set(message)
        anchor_status_hold["until"] = time.time() + hold
        alignment_status_cache["message"] = None

    refresh_active_host_label()
    refresh_active_model_label()

    # Scrollable container
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
        step = int(-1 * (event.delta / 120))
        canvas.yview_scroll(step, "units")

    def _on_mousewheel_linux(event, direction: int):
        canvas.yview_scroll(direction, "units")

    main.bind("<Configure>", _sync_scroll_region)
    canvas.bind("<Configure>", _sync_scroll_region)
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    canvas.bind_all("<Button-4>", lambda e: _on_mousewheel_linux(e, -1))
    canvas.bind_all("<Button-5>", lambda e: _on_mousewheel_linux(e, 1))

    # --- Capture Region ---
    frm_region = ttk.LabelFrame(main, text="Capture Region", style="Glass.TLabelframe")
    frm_region.pack(fill="x", padx=5, pady=8)

    slider_left = create_glass_scale(
        frm_region, text="Left", minimum=0, maximum=3000,
        initial=app_state.cap_region["left"], command=update_region_from_sliders,
    )
    slider_top = create_glass_scale(
        frm_region, text="Top", minimum=0, maximum=2000,
        initial=app_state.cap_region["top"], command=update_region_from_sliders,
    )
    slider_width = create_glass_scale(
        frm_region, text="Width", minimum=50, maximum=1000,
        initial=app_state.cap_region["width"], command=update_region_from_sliders,
    )
    slider_height = create_glass_scale(
        frm_region, text="Height", minimum=20, maximum=500,
        initial=app_state.cap_region["height"], command=update_region_from_sliders,
        padding=(0, 0),
    )

    register_capture_sliders(slider_left, slider_top, slider_width, slider_height)
    sync_capture_sliders()

    # --- Head Sway Compensation ---
    frm_anchor = ttk.LabelFrame(main, text="Head Sway Compensation", style="Glass.TLabelframe")
    frm_anchor.pack(fill="x", padx=5, pady=8)

    auto_align_var = tk.BooleanVar(value=app_state.auto_align_enabled)
    chk_auto_align = ttk.Checkbutton(
        frm_anchor, text="Enable auto alignment",
        variable=auto_align_var, command=toggle_auto_align,
        style="Glass.TCheckbutton",
    )
    chk_auto_align.pack(anchor="w", padx=5, pady=(5, 0))

    anchor_overlay_var = tk.BooleanVar(value=app_state.anchor_overlay_visible)
    chk_anchor_overlay = ttk.Checkbutton(
        frm_anchor, text="Show anchor overlay",
        variable=anchor_overlay_var, command=toggle_anchor_overlay_visibility,
        style="Glass.TCheckbutton",
    )
    chk_anchor_overlay.pack(anchor="w", padx=5, pady=(0, 5))

    interval_row = ttk.Frame(frm_anchor, style="Glass.Section.TFrame")
    interval_row.pack(fill="x", padx=5, pady=(0, 5))
    ttk.Label(interval_row, text="Alignment interval (ms)", style="Glass.Small.TLabel").pack(side="left")
    alignment_interval_var = tk.IntVar(value=int(app_state.alignment_poll_interval_ms))
    alignment_interval_spin = tk.Spinbox(
        interval_row, from_=100, to=5000, increment=50,
        textvariable=alignment_interval_var, width=6,
        command=update_alignment_interval,
    )
    alignment_interval_spin.pack(side="left", padx=5)
    style_spinbox(alignment_interval_spin, colors)
    alignment_interval_var.trace_add("write", update_alignment_interval)

    threshold_row = ttk.Frame(frm_anchor, style="Glass.Section.TFrame")
    threshold_row.pack(fill="x", padx=5, pady=5)
    ttk.Label(threshold_row, text="Detection threshold", style="Glass.Small.TLabel").pack(side="left")
    threshold_var = tk.DoubleVar(value=app_state.anchor_threshold)
    threshold_spin = tk.Spinbox(
        threshold_row, from_=0.10, to=0.99, increment=0.01,
        textvariable=threshold_var, width=6, command=update_threshold,
    )
    threshold_spin.pack(side="left", padx=5)
    style_spinbox(threshold_spin, colors)
    threshold_var.trace_add("write", update_threshold)

    anchor_left = create_glass_scale(
        frm_anchor, text="Anchor Left", minimum=0, maximum=3840,
        initial=app_state.anchor_region["left"], command=update_anchor_region_from_sliders,
    )
    anchor_top = create_glass_scale(
        frm_anchor, text="Anchor Top", minimum=0, maximum=2160,
        initial=app_state.anchor_region["top"], command=update_anchor_region_from_sliders,
    )
    anchor_width = create_glass_scale(
        frm_anchor, text="Anchor Width", minimum=50, maximum=1200,
        initial=app_state.anchor_region["width"], command=update_anchor_region_from_sliders,
    )
    anchor_height = create_glass_scale(
        frm_anchor, text="Anchor Height", minimum=50, maximum=800,
        initial=app_state.anchor_region["height"], command=update_anchor_region_from_sliders,
    )
    anchor_offset_x = create_glass_scale(
        frm_anchor, text="Offset X", minimum=-300, maximum=600,
        initial=app_state.anchor_offset["x"], command=update_anchor_offset_from_sliders,
    )
    anchor_offset_y = create_glass_scale(
        frm_anchor, text="Offset Y", minimum=-300, maximum=600,
        initial=app_state.anchor_offset["y"], command=update_anchor_offset_from_sliders,
        padding=(0, 0),
    )

    register_anchor_sliders(anchor_left, anchor_top, anchor_width, anchor_height, anchor_offset_x, anchor_offset_y)
    sync_anchor_sliders()

    # --- Ollama Connection ---
    frm_network = ttk.LabelFrame(main, text="Ollama Connection", style="Glass.TLabelframe")
    frm_network.pack(fill="x", padx=5, pady=8)
    ttk.Label(
        frm_network,
        text="Ollama model (set in config.json or environment).",
        style="Glass.Small.TLabel", wraplength=360, justify="left",
    ).pack(fill="x", padx=5, pady=(5, 2))
    model_entry = ttk.Combobox(
        frm_network, textvariable=ollama_model_var,
        values=[
            "moondream:1.8b",  # Ultra-light, edge-optimized
            "granite3.2-vision:2b",  # IBM, compact, UI/text focused
            "deepseek-ocr:3b",  # Dedicated OCR, accurate
            "smolvlm",  # Hugging Face, ultra-small
            "bakllava:1.8b",  # Another compact VLM
            "llava:1.5b",  # Small LLaVA variant
            "qwen2.5vl:3b", "qwen3-vl:2b", "qwen3-vl:4b"  # Qwen family for reference
        ],
    )
    model_entry.pack(fill="x", padx=5, pady=(0, 5))
    ttk.Label(
        frm_network,
        text="Remote Ollama host (IPv4/hostname with optional port). Leave blank to use this PC.",
        style="Glass.Small.TLabel", wraplength=360, justify="left",
    ).pack(fill="x", padx=5, pady=(5, 2))
    host_entry = ttk.Entry(frm_network, textvariable=ollama_host_var)
    host_entry.pack(fill="x", padx=5, pady=(0, 5))

    network_button_row = ttk.Frame(frm_network, style="Glass.Section.TFrame")
    network_button_row.pack(fill="x", padx=5, pady=(0, 5))
    ttk.Button(network_button_row, text="Apply Host", command=apply_ollama_host_from_ui, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(network_button_row, text="Apply Model", command=apply_ollama_model_from_ui, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(network_button_row, text="Use Localhost", command=use_local_ollama_host, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(network_button_row, text="Open Mobile UI", command=open_mobile_overlay, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Label(frm_network, textvariable=ollama_active_host_var, style="Glass.Small.TLabel", justify="left").pack(fill="x", padx=5, pady=(0, 2))
    ttk.Label(frm_network, textvariable=ollama_active_model_var, style="Glass.Small.TLabel", justify="left").pack(fill="x", padx=5, pady=(0, 5))

    # Anchor buttons (placed after network section, inside frm_anchor)
    anchor_btn_row = ttk.Frame(frm_anchor, style="Glass.Section.TFrame")
    anchor_btn_row.pack(fill="x", padx=5, pady=5)
    ttk.Button(anchor_btn_row, text="Reload Templates", command=reload_anchor_templates, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(anchor_btn_row, text="Realign Now", command=manual_realign, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(anchor_btn_row, text="Open Template Folder", command=open_anchor_directory, style="Glass.TButton").pack(side="left", padx=5)

    # --- Result Display ---
    frm_display = ttk.LabelFrame(main, text="Result Display", style="Glass.TLabelframe")
    frm_display.pack(fill="x", padx=5, pady=8)

    info_offset_x = create_glass_scale(
        frm_display, text="Display offset X", minimum=-800, maximum=800,
        initial=int(app_state.info_overlay_offset.get("x", 0)),
        command=update_info_overlay_from_sliders,
    )
    info_offset_y = create_glass_scale(
        frm_display, text="Display offset Y", minimum=-600, maximum=600,
        initial=int(app_state.info_overlay_offset.get("y", 0)),
        command=update_info_overlay_from_sliders,
        padding=(0, 0),
    )

    register_overlay_sliders(info_offset_x, info_offset_y)
    sync_overlay_sliders()

    # --- Controls ---
    frm_ctrl = ttk.LabelFrame(main, text="Controls", style="Glass.TLabelframe")
    frm_ctrl.pack(fill="x", padx=5, pady=8)

    capture_interval_frame = ttk.Frame(frm_ctrl, style="Glass.Section.TFrame")
    capture_interval_frame.pack(fill="x", padx=5, pady=(5, 10))
    ttk.Label(capture_interval_frame, text="Continuous capture interval (s)", style="Glass.Small.TLabel").pack(side="left")
    capture_interval_var = tk.DoubleVar(value=float(app_state.continuous_capture_interval))
    capture_interval_spin = tk.Spinbox(
        capture_interval_frame, from_=0.2, to=30.0, increment=0.1,
        textvariable=capture_interval_var, width=6, format="%.1f",
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
