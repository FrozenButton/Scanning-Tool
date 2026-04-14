"""Slider registration and synchronization helpers for overlays."""

import logging
import tkinter as tk
from typing import Dict

from scanning_tool.state import ScaleWidget, app_state
from .base import safe_tk

logger = logging.getLogger("scanning_tool")


def register_capture_sliders(left: ScaleWidget, top: ScaleWidget, width: ScaleWidget, height: ScaleWidget) -> None:
    app_state.gui_control_state["capture"].update({
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
    app_state.gui_control_state["anchor"].update({
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "offset_x": offset_x,
        "offset_y": offset_y,
    })


def register_overlay_sliders(offset_x: ScaleWidget, offset_y: ScaleWidget) -> None:
    app_state.gui_control_state["overlay"].update({"offset_x": offset_x, "offset_y": offset_y})


def sync_capture_sliders() -> None:
    state = app_state.gui_control_state
    widgets = state["capture"]
    widget = widgets["left"]
    if not widget or state["syncing"]["capture"]:
        return

    def _apply() -> None:
        if state["syncing"]["capture"]:
            return
        state["syncing"]["capture"] = True
        try:
            try:
                widgets["left"].set(int(app_state.cap_region["left"]))
                widgets["top"].set(int(app_state.cap_region["top"]))
                widgets["width"].set(int(app_state.cap_region["width"]))
                widgets["height"].set(int(app_state.cap_region["height"]))
            except tk.TclError:
                pass
        finally:
            state["syncing"]["capture"] = False

    safe_tk(lambda: widget.after(0, _apply))


def sync_anchor_sliders() -> None:
    state = app_state.gui_control_state
    widgets = state["anchor"]
    widget = widgets["left"]
    if not widget or state["syncing"]["anchor"]:
        return

    def _apply() -> None:
        if state["syncing"]["anchor"]:
            return
        state["syncing"]["anchor"] = True
        try:
            try:
                widgets["left"].set(int(app_state.anchor_region["left"]))
                widgets["top"].set(int(app_state.anchor_region["top"]))
                widgets["width"].set(int(app_state.anchor_region["width"]))
                widgets["height"].set(int(app_state.anchor_region["height"]))
                widgets["offset_x"].set(int(app_state.anchor_offset["x"]))
                widgets["offset_y"].set(int(app_state.anchor_offset["y"]))
            except tk.TclError:
                pass
        finally:
            state["syncing"]["anchor"] = False

    safe_tk(lambda: widget.after(0, _apply))


def sync_overlay_sliders() -> None:
    state = app_state.gui_control_state
    widgets = state["overlay"]
    widget = widgets["offset_x"]
    if not widget or state["syncing"]["overlay"]:
        return

    def _apply() -> None:
        if state["syncing"]["overlay"]:
            return
        state["syncing"]["overlay"] = True
        try:
            try:
                widgets["offset_x"].set(int(app_state.info_overlay_offset.get("x", 0)))
                widgets["offset_y"].set(int(app_state.info_overlay_offset.get("y", 0)))
            except tk.TclError:
                pass
        finally:
            state["syncing"]["overlay"] = False

    safe_tk(lambda: widget.after(0, _apply))
