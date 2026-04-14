"""Ollama Connection section — model picker, host entry, and action buttons."""

import logging
import webbrowser

import tkinter as tk
from tkinter import ttk

from scanning_tool.gui.sections.base import SectionContext
from scanning_tool.gui.widgets import (
    create_button_row,
    create_labeled_combobox,
    create_labeled_entry,
)
from scanning_tool.ollama import (
    ensure_model_installed,
    get_ollama_host,
    get_ollama_model,
    log_model_running_status,
    set_configured_ollama_host,
    set_configured_ollama_model,
)
from scanning_tool.state_context import app_state
from scanning_tool.web import get_local_ip

logger = logging.getLogger("scanning_tool")


SUGGESTED_MODELS = (
    "moondream:1.8b",
    "granite3.2-vision:2b",
    "deepseek-ocr:3b",
    "smolvlm",
    "bakllava:1.8b",
    "llava:1.5b",
    "qwen2.5vl:3b",
    "qwen3-vl:2b",
    "qwen3-vl:4b",
)


class OllamaSection:
    """Model selector, host entry, and active-status labels for Ollama."""

    def build(self, parent: ttk.Widget, ctx: SectionContext) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text="Ollama Connection", style="Glass.TLabelframe")
        frame.pack(fill="x", padx=5, pady=8)

        self._status = ctx.status
        self._host_var = tk.StringVar(value=app_state.settings.ollama.configured_ollama_host)
        self._model_var = tk.StringVar(value=app_state.settings.ollama.ollama_model)
        self._active_host_var = tk.StringVar()
        self._active_model_var = tk.StringVar()

        create_labeled_combobox(
            frame,
            text="Ollama model (set in config.json or environment).",
            variable=self._model_var,
            values=list(SUGGESTED_MODELS),
            width=48,
        )

        create_labeled_entry(
            frame,
            text="Remote Ollama host (IPv4/hostname with optional port). Leave blank to use this PC.",
            variable=self._host_var,
            colors=ctx.colors,
        )

        create_button_row(
            frame,
            [
                ("Apply Host", self._apply_host),
                ("Apply Model", self._apply_model),
                ("Use Localhost", self._use_localhost),
                ("Open Mobile UI", self._open_mobile_overlay),
            ],
        )

        ttk.Label(
            frame, textvariable=self._active_host_var,
            style="Glass.Small.TLabel", justify="left",
        ).pack(fill="x", padx=5, pady=(0, 2))
        ttk.Label(
            frame, textvariable=self._active_model_var,
            style="Glass.Small.TLabel", justify="left",
        ).pack(fill="x", padx=5, pady=(0, 5))

        self._refresh_active_labels()
        return frame

    def _refresh_active_labels(self) -> None:
        self._active_host_var.set(f"Active host: {get_ollama_host()}")
        self._active_model_var.set(f"Active model: {get_ollama_model()}")

    def _apply_model(self) -> None:
        model_value = self._model_var.get().strip()
        if not model_value:
            self._status.set_status("Please specify an Ollama model.")
            return
        set_configured_ollama_model(model_value)
        try:
            ensure_model_installed(model_value, exit_on_error=False)
        except Exception as exc:
            self._status.set_status(f"Model install failed: {exc}")
            logger.error("Failed to install model %s: %s", model_value, exc)
            return
        running = log_model_running_status(model_value)
        self._refresh_active_labels()
        if running:
            self._status.set_status(
                f"Ollama model set to {model_value} and is currently running."
            )
        else:
            self._status.set_status(
                f"Ollama model set to {model_value}. It is not running yet and will start on first scan."
            )
        logger.info("Ollama model set to %s.", model_value)

    def _apply_host(self) -> None:
        sanitized = set_configured_ollama_host(self._host_var.get())
        self._refresh_active_labels()
        active_host = get_ollama_host()
        if sanitized:
            self._host_var.set(sanitized)
            message = f"Remote Ollama host set to {active_host}."
        else:
            message = f"Ollama host cleared. Using {active_host}."
        self._status.set_status(message)
        logger.info(message)

    def _use_localhost(self) -> None:
        self._host_var.set("")
        set_configured_ollama_host("")
        self._refresh_active_labels()
        message = f"Ollama host cleared. Using {get_ollama_host()}."
        self._status.set_status(message)
        logger.info(message)

    def _open_mobile_overlay(self) -> None:
        url = f"http://{get_local_ip()}:5000"
        try:
            webbrowser.open_new_tab(url)
        except Exception as exc:
            self._status.set_status(f"Unable to open browser: {exc}")
            logger.warning("Failed to open mobile overlay URL %s: %s", url, exc)
        else:
            self._status.set_status(f"Opening overlay in browser: {url}")
            logger.info("Opened mobile overlay URL: %s", url)
