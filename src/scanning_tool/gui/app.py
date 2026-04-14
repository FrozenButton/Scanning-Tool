"""Main application entry: assembles the Tk root and all GUI sections."""

import logging
import os
import subprocess
import sys
import webbrowser

import tkinter as tk
from tkinter import ttk

from scanning_tool.anchor import AnchorRegionTracker, perform_auto_alignment
from scanning_tool.config import ensure_anchor_directory
from scanning_tool.gui.alignment import AlignmentPoller
from scanning_tool.gui.lifecycle import register_close_handler
from scanning_tool.gui.sections import (
    CaptureRegionSection,
    ControlsSection,
    ResultDisplaySection,
    SectionContext,
)
from scanning_tool.gui.status import StatusBar
from scanning_tool.gui.theme import apply_glass_theme, style_spinbox
from scanning_tool.gui.widgets import ScrollableFrame, create_glass_scale
from scanning_tool.ollama_service import (
    ensure_model_installed,
    get_ollama_host,
    get_ollama_model,
    log_model_running_status,
    set_configured_ollama_host,
    set_configured_ollama_model,
)
from scanning_tool.overlay import (
    hide_anchor_overlay,
    register_anchor_sliders,
    show_anchor_overlay,
    show_overlay,
    sync_anchor_sliders,
    update_anchor_overlay_region,
)
from scanning_tool.state import app_state
from scanning_tool.web import get_local_ip

logger = logging.getLogger("scanning_tool")


def launch_gui() -> None:
    """Build and run the main Tkinter control panel."""
    root = tk.Tk()
    root.title("Star Citizen Scanner Control")
    register_close_handler(root)

    colors = apply_glass_theme(root)
    status = StatusBar(root)
    status.install_as_scanning_callback()
    ctx = SectionContext(root=root, colors=colors, status=status)

    ollama_host_var = tk.StringVar(value=app_state.configured_ollama_host)
    ollama_model_var = tk.StringVar(value=app_state.ollama_model)
    ollama_active_host_var = tk.StringVar()
    ollama_active_model_var = tk.StringVar(value=f"Active model: {get_ollama_model()}")

    def refresh_active_host_label() -> None:
        ollama_active_host_var.set(f"Active host: {get_ollama_host()}")

    def refresh_active_model_label() -> None:
        ollama_active_model_var.set(f"Active model: {get_ollama_model()}")

    def update_anchor_region_from_sliders(*_args: object) -> None:
        if app_state.gui_control_state["syncing"]["anchor"]:
            return
        app_state.anchor_region["left"] = int(anchor_left.get())
        app_state.anchor_region["top"] = int(anchor_top.get())
        app_state.anchor_region["width"] = int(anchor_width.get())
        app_state.anchor_region["height"] = int(anchor_height.get())
        status.set_anchor(f"Anchor region updated: {app_state.anchor_region}", hold=2.0)
        if app_state.auto_align_enabled:
            perform_auto_alignment()
        update_anchor_overlay_region()

    def update_anchor_offset_from_sliders(*_args: object) -> None:
        if app_state.gui_control_state["syncing"]["anchor"]:
            return
        app_state.anchor_offset["x"] = int(anchor_offset_x.get())
        app_state.anchor_offset["y"] = int(anchor_offset_y.get())
        status.set_anchor(f"Anchor offset updated: {app_state.anchor_offset}", hold=2.0)
        if app_state.auto_align_enabled:
            perform_auto_alignment()

    def toggle_auto_align() -> None:
        app_state.auto_align_enabled = auto_align_var.get()
        app_state.last_alignment_info["enabled"] = app_state.auto_align_enabled
        if app_state.auto_align_enabled:
            status.set_anchor("Head sway compensation enabled.")
            perform_auto_alignment()
        else:
            status.set_anchor("Head sway compensation disabled.")

    def reload_anchor_templates() -> None:
        ensure_anchor_directory(app_state.anchor_template_dir)
        if app_state.anchor_tracker is None:
            app_state.anchor_tracker = AnchorRegionTracker(
                app_state.anchor_template_dir, app_state.anchor_threshold
            )
        count = app_state.anchor_tracker.set_directory(app_state.anchor_template_dir)
        status.set_anchor(
            f"Loaded {count} anchor template(s) from {app_state.anchor_template_dir}."
        )

    def manual_realign() -> None:
        success = perform_auto_alignment()
        if success:
            status.set_anchor(
                f"Anchor locked using {app_state.last_alignment_info['template']} "
                f"(score {app_state.last_alignment_info['score']:.2f}).",
                hold=2.5,
            )
            status.set_status(f"Auto alignment adjusted CAP_REGION: {app_state.cap_region}")
        else:
            status.set_anchor("Anchor match not found. Adjust search region or add templates.")

    def open_anchor_directory() -> None:
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
            status.set_anchor(f"Unable to open template folder: {exc}", hold=3.0)
        else:
            status.set_anchor(f"Opened template folder: {path}")

    def update_threshold(*_args: object) -> None:
        try:
            value = float(threshold_var.get())
        except (tk.TclError, ValueError):
            return
        value = max(0.1, min(0.99, value))
        app_state.anchor_threshold = value
        if app_state.anchor_tracker is not None:
            app_state.anchor_tracker.set_threshold(app_state.anchor_threshold)
        status.set_anchor(f"Anchor detection threshold set to {app_state.anchor_threshold:.2f}")

    def toggle_anchor_overlay_visibility() -> None:
        app_state.anchor_overlay_visible = anchor_overlay_var.get()
        if app_state.anchor_overlay_visible:
            show_anchor_overlay()
            status.set_anchor("Anchor overlay shown.")
        else:
            hide_anchor_overlay()
            status.set_anchor("Anchor overlay hidden.")

    def update_alignment_interval(*_args: object) -> None:
        try:
            value = int(alignment_interval_var.get())
        except (tk.TclError, ValueError):
            return
        value = max(100, min(5000, value))
        app_state.alignment_poll_interval_ms = value
        status.set_anchor(
            f"Alignment interval set to {app_state.alignment_poll_interval_ms} ms", hold=2.0
        )

    def apply_ollama_model_from_ui() -> None:
        model_value = ollama_model_var.get().strip()
        if not model_value:
            status.set_status("Please specify an Ollama model.")
            return
        set_configured_ollama_model(model_value)
        try:
            ensure_model_installed(model_value, exit_on_error=False)
        except Exception as exc:
            status.set_status(f"Model install failed: {exc}")
            logger.error("Failed to install model %s: %s", model_value, exc)
            return
        running = log_model_running_status(model_value)
        refresh_active_model_label()
        if running:
            status.set_status(f"Ollama model set to {model_value} and is currently running.")
        else:
            status.set_status(
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
        status.set_status(message)
        logger.info(message)

    def use_local_ollama_host() -> None:
        ollama_host_var.set("")
        set_configured_ollama_host("")
        refresh_active_host_label()
        active_host = get_ollama_host()
        message = f"Ollama host cleared. Using {active_host}."
        status.set_status(message)
        logger.info(message)

    def open_mobile_overlay() -> None:
        url = f"http://{get_local_ip()}:5000"
        try:
            webbrowser.open_new_tab(url)
        except Exception as exc:
            status.set_status(f"Unable to open browser: {exc}")
            logger.warning("Failed to open mobile overlay URL %s: %s", url, exc)
        else:
            status.set_status(f"Opening overlay in browser: {url}")
            logger.info("Opened mobile overlay URL: %s", url)

    refresh_active_host_label()
    refresh_active_model_label()

    scroll = ScrollableFrame(root, colors)
    main = scroll.inner

    CaptureRegionSection().build(main, ctx)

    # --- Head Sway Compensation ---
    frm_anchor = ttk.LabelFrame(main, text="Head Sway Compensation", style="Glass.TLabelframe")
    frm_anchor.pack(fill="x", padx=5, pady=8)

    auto_align_var = tk.BooleanVar(value=app_state.auto_align_enabled)
    ttk.Checkbutton(
        frm_anchor, text="Enable auto alignment",
        variable=auto_align_var, command=toggle_auto_align,
        style="Glass.TCheckbutton",
    ).pack(anchor="w", padx=5, pady=(5, 0))

    anchor_overlay_var = tk.BooleanVar(value=app_state.anchor_overlay_visible)
    ttk.Checkbutton(
        frm_anchor, text="Show anchor overlay",
        variable=anchor_overlay_var, command=toggle_anchor_overlay_visibility,
        style="Glass.TCheckbutton",
    ).pack(anchor="w", padx=5, pady=(0, 5))

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

    register_anchor_sliders(
        anchor_left, anchor_top, anchor_width, anchor_height, anchor_offset_x, anchor_offset_y
    )
    sync_anchor_sliders()

    # --- Ollama Connection ---
    frm_network = ttk.LabelFrame(main, text="Ollama Connection", style="Glass.TLabelframe")
    frm_network.pack(fill="x", padx=5, pady=8)
    ttk.Label(
        frm_network,
        text="Ollama model (set in config.json or environment).",
        style="Glass.Small.TLabel", wraplength=360, justify="left",
    ).pack(fill="x", padx=5, pady=(5, 2))
    ttk.Combobox(
        frm_network, textvariable=ollama_model_var,
        values=[
            "moondream:1.8b",
            "granite3.2-vision:2b",
            "deepseek-ocr:3b",
            "smolvlm",
            "bakllava:1.8b",
            "llava:1.5b",
            "qwen2.5vl:3b", "qwen3-vl:2b", "qwen3-vl:4b",
        ],
    ).pack(fill="x", padx=5, pady=(0, 5))
    ttk.Label(
        frm_network,
        text="Remote Ollama host (IPv4/hostname with optional port). Leave blank to use this PC.",
        style="Glass.Small.TLabel", wraplength=360, justify="left",
    ).pack(fill="x", padx=5, pady=(5, 2))
    ttk.Entry(frm_network, textvariable=ollama_host_var).pack(fill="x", padx=5, pady=(0, 5))

    network_button_row = ttk.Frame(frm_network, style="Glass.Section.TFrame")
    network_button_row.pack(fill="x", padx=5, pady=(0, 5))
    ttk.Button(network_button_row, text="Apply Host", command=apply_ollama_host_from_ui, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(network_button_row, text="Apply Model", command=apply_ollama_model_from_ui, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(network_button_row, text="Use Localhost", command=use_local_ollama_host, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(network_button_row, text="Open Mobile UI", command=open_mobile_overlay, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Label(frm_network, textvariable=ollama_active_host_var, style="Glass.Small.TLabel", justify="left").pack(fill="x", padx=5, pady=(0, 2))
    ttk.Label(frm_network, textvariable=ollama_active_model_var, style="Glass.Small.TLabel", justify="left").pack(fill="x", padx=5, pady=(0, 5))

    anchor_btn_row = ttk.Frame(frm_anchor, style="Glass.Section.TFrame")
    anchor_btn_row.pack(fill="x", padx=5, pady=5)
    ttk.Button(anchor_btn_row, text="Reload Templates", command=reload_anchor_templates, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(anchor_btn_row, text="Realign Now", command=manual_realign, style="Glass.TButton").pack(side="left", padx=5)
    ttk.Button(anchor_btn_row, text="Open Template Folder", command=open_anchor_directory, style="Glass.TButton").pack(side="left", padx=5)

    ResultDisplaySection().build(main, ctx)
    ControlsSection().build(main, ctx)

    ttk.Label(main, textvariable=status.status_var, anchor="w", justify="left", style="Glass.Status.TLabel").pack(
        fill="x", padx=5, pady=(8, 0)
    )
    ttk.Label(main, textvariable=status.anchor_status_var, anchor="w", justify="left", style="Glass.Subtle.TLabel").pack(
        fill="x", padx=5, pady=(2, 5)
    )

    root.update_idletasks()
    show_overlay(root.winfo_screenwidth(), root.winfo_screenheight())
    AlignmentPoller(root, status).start()
    root.mainloop()
