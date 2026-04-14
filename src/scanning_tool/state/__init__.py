"""State package API for the scanning tool."""

from __future__ import annotations

from typing import Any, Dict, Tuple, Union

try:
    import tkinter as tk
    from tkinter import ttk
except ImportError:  # headless environments
    tk = None  # type: ignore[assignment]
    ttk = None  # type: ignore[assignment]

from .context import AppContext, app_state as _context_app_state

ScaleWidget = Union[tk.Scale, ttk.Scale]

_ATTRIBUTE_PATHS: Dict[str, Tuple[str, ...]] = {
    # Persisted settings
    "cap_region": ("settings", "capture", "cap_region"),
    "anchor_region": ("settings", "anchor", "anchor_region"),
    "anchor_offset": ("settings", "anchor", "anchor_offset"),
    "anchor_threshold": ("settings", "anchor", "anchor_threshold"),
    "auto_align_enabled": ("settings", "anchor", "auto_align_enabled"),
    "anchor_template_dir": ("settings", "anchor", "anchor_template_dir"),
    "alignment_poll_interval_ms": ("settings", "anchor", "alignment_poll_interval_ms"),
    "continuous_capture_interval": ("settings", "capture", "continuous_capture_interval"),
    "info_overlay_offset": ("settings", "overlay", "info_overlay_offset"),
    "label_color": ("settings", "overlay", "label_color"),
    "ollama_model": ("settings", "ollama", "ollama_model"),
    "configured_ollama_host": ("settings", "ollama", "configured_ollama_host"),
    "default_ollama_host": ("settings", "ollama", "default_ollama_host"),
    "min_confidence": ("settings", "scan", "min_confidence"),
    "debug_show_overlay": ("settings", "overlay", "debug_show_overlay"),
    # Runtime scan state
    "last_result": ("scan_state", "last_result"),
    "last_alignment_info": ("scan_state", "last_alignment_info"),
    "continuous_mode": ("scan_state", "continuous_mode"),
    "anchor_tracker": ("scan_state", "anchor_tracker"),
    # Runtime service state
    "ollama_client": ("service_state", "ollama_client"),
    "ollama_client_host": ("service_state", "ollama_client_host"),
    "ollama_server_process": ("service_state", "ollama_server_process"),
    "code_re": ("service_state", "code_re"),
    "host_scheme_re": ("service_state", "host_scheme_re"),
    "rock_data": ("service_state", "rock_data"),
    "deposit_tables": ("service_state", "deposit_tables"),
    "gui_status_callback": ("service_state", "gui_status_callback"),
    # Overlay state
    "capture_overlay_root": ("overlay_state", "capture_overlay_root"),
    "capture_overlay_canvas": ("overlay_state", "capture_overlay_canvas"),
    "capture_rect_id": ("overlay_state", "capture_rect_id"),
    "capture_overlay_animation_job": ("overlay_state", "capture_overlay_animation_job"),
    "capture_overlay_last_layout": ("overlay_state", "capture_overlay_last_layout"),
    "info_overlay_root": ("overlay_state", "info_overlay_root"),
    "info_overlay_canvas": ("overlay_state", "info_overlay_canvas"),
    "info_text_id": ("overlay_state", "info_text_id"),
    "info_overlay_geometry": ("overlay_state", "info_overlay_geometry"),
    "overlay_text": ("overlay_state", "overlay_text"),
    "last_overlay_time": ("overlay_state", "last_overlay_time"),
    "anchor_overlay_root": ("overlay_state", "anchor_overlay_root"),
    "anchor_overlay_canvas": ("overlay_state", "anchor_overlay_canvas"),
    "anchor_rect_id": ("overlay_state", "anchor_rect_id"),
    "border_canvas": ("overlay_state", "border_canvas"),
    "show_border": ("overlay_state", "show_border"),
    "anchor_overlay_visible": ("overlay_state", "anchor_overlay_visible"),
    # Control state
    "gui_control_state": ("control_state", "gui_control_state"),
}


def _resolve_path(path: Tuple[str, ...]) -> Any:
    node: object = _context_app_state
    for key in path:
        if isinstance(node, dict):
            node = node[key]
        else:
            node = getattr(node, key)
    return node


def _assign_path(path: Tuple[str, ...], value: Any) -> None:
    node: object = _context_app_state
    for key in path[:-1]:
        if isinstance(node, dict):
            node = node[key]
        else:
            node = getattr(node, key)
    final_key = path[-1]
    if isinstance(node, dict):
        node[final_key] = value
    else:
        setattr(node, final_key, value)


class AppStateShim:
    def __getattr__(self, name: str) -> Any:
        if name in {"settings", "scan_state", "overlay_state", "control_state", "service_state"}:
            return getattr(_context_app_state, name)
        path = _ATTRIBUTE_PATHS.get(name)
        if path is not None:
            return _resolve_path(path)
        raise AttributeError(f"AppState has no attribute {name!r}")

    def __setattr__(self, name: str, value: Any) -> None:
        if name == "_context_app_state":
            object.__setattr__(self, name, value)
            return
        if name in {"settings", "scan_state", "overlay_state", "control_state", "service_state"}:
            raise AttributeError(f"Cannot reassign {name}")
        path = _ATTRIBUTE_PATHS.get(name)
        if path is not None:
            _assign_path(path, value)
            return
        raise AttributeError(f"AppState has no attribute {name!r}")

    def __dir__(self) -> list[str]:
        return sorted(super().__dir__() + list(_ATTRIBUTE_PATHS) + [
            "settings", "scan_state", "overlay_state", "control_state", "service_state"
        ])


app_state = AppStateShim()

__all__ = ["AppContext", "app_state", "ScaleWidget"]
