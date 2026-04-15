"""Slider registration and synchronization helpers for overlays."""

from loguru import logger
import tkinter as tk
from typing import Dict

from scanning_tool.gui.control_state import ScaleWidget
from scanning_tool.state import app_state
from .base import safe_tk



def register_capture_sliders(left: ScaleWidget, top: ScaleWidget, width: ScaleWidget, height: ScaleWidget) -> None:
    app_state.control_state.gui_control_state["capture"].update({
        "left": left, "top": top, "width": width, "height": height,
    })


def register_anchor_sliders(
    left: ScaleWidget,
    top: ScaleWidget,
    width: ScaleWidget,
    height: ScaleWidget,
    offset_x: ScaleWidget,
    offset_y: ScaleWidget,
) -> None:
    app_state.control_state.gui_control_state["anchor"].update({
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "offset_x": offset_x,
        "offset_y": offset_y,
    })


def register_overlay_sliders(offset_x: ScaleWidget, offset_y: ScaleWidget) -> None:
    app_state.control_state.gui_control_state["overlay"].update({"offset_x": offset_x, "offset_y": offset_y})


def sync_capture_sliders() -> None:
    state = app_state.control_state.gui_control_state
    widgets = state["capture"]
    widget = widgets["left"]
    if not widget or state["syncing"]["capture"]:
        return

    capture_region = app_state.settings.capture.cap_region

    def _apply() -> None:
        if state["syncing"]["capture"]:
            return
        state["syncing"]["capture"] = True
        try:
            try:
                widgets["left"].set(int(capture_region["left"]))
                widgets["top"].set(int(capture_region["top"]))
                widgets["width"].set(int(capture_region["width"]))
                widgets["height"].set(int(capture_region["height"]))
            except tk.TclError:
                pass
        finally:
            state["syncing"]["capture"] = False

    safe_tk(lambda: widget.after(0, _apply))


def sync_anchor_sliders() -> None:
    state = app_state.control_state.gui_control_state
    widgets = state["anchor"]
    widget = widgets["left"]
    if not widget or state["syncing"]["anchor"]:
        return

    anchor_region = app_state.settings.anchor.anchor_region
    anchor_offset = app_state.settings.anchor.anchor_offset

    def _apply() -> None:
        if state["syncing"]["anchor"]:
            return
        state["syncing"]["anchor"] = True
        try:
            try:
                widgets["left"].set(int(anchor_region["left"]))
                widgets["top"].set(int(anchor_region["top"]))
                widgets["width"].set(int(anchor_region["width"]))
                widgets["height"].set(int(anchor_region["height"]))
                widgets["offset_x"].set(int(anchor_offset["x"]))
                widgets["offset_y"].set(int(anchor_offset["y"]))
            except tk.TclError:
                pass
        finally:
            state["syncing"]["anchor"] = False

    safe_tk(lambda: widget.after(0, _apply))


def sync_overlay_sliders() -> None:
    state = app_state.control_state.gui_control_state
    widgets = state["overlay"]
    widget = widgets["offset_x"]
    if not widget or state["syncing"]["overlay"]:
        return

    overlay_offset = app_state.settings.overlay.info_overlay_offset

    def _apply() -> None:
        if state["syncing"]["overlay"]:
            return
        state["syncing"]["overlay"] = True
        try:
            try:
                widgets["offset_x"].set(int(overlay_offset.get("x", 0)))
                widgets["offset_y"].set(int(overlay_offset.get("y", 0)))
            except tk.TclError:
                pass
        finally:
            state["syncing"]["overlay"] = False

    safe_tk(lambda: widget.after(0, _apply))
