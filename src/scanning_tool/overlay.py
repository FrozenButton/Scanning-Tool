"""Overlay window management: capture, info, and anchor overlays."""

import logging
import time
from typing import Dict, Optional, Tuple, TypedDict


# Strongly-typed info dict for overlay label
class OverlayInfo(TypedDict, total=False):
    name: str
    deposits: int


import tkinter as tk

from scanning_tool.state import app_state, ScaleWidget

logger = logging.getLogger("scanning_tool")


# ---------- Slider registration and sync ----------

def register_capture_sliders(left: ScaleWidget, top: ScaleWidget, width: ScaleWidget, height: ScaleWidget) -> None:
    app_state.gui_control_state["capture"].update({
        "left": left, "top": top, "width": width, "height": height,
    })


def register_anchor_sliders(
    left: ScaleWidget, top: ScaleWidget, width: ScaleWidget, height: ScaleWidget,
    offset_x: ScaleWidget, offset_y: ScaleWidget,
) -> None:
    app_state.gui_control_state["anchor"].update({
        "left": left, "top": top, "width": width, "height": height,
        "offset_x": offset_x, "offset_y": offset_y,
    })


def register_overlay_sliders(offset_x: ScaleWidget, offset_y: ScaleWidget) -> None:
    app_state.gui_control_state["overlay"].update({"offset_x": offset_x, "offset_y": offset_y})


def sync_capture_sliders() -> None:
    state = app_state.gui_control_state
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
                widgets["left"].set(int(app_state.cap_region["left"]))
                widgets["top"].set(int(app_state.cap_region["top"]))
                widgets["width"].set(int(app_state.cap_region["width"]))
                widgets["height"].set(int(app_state.cap_region["height"]))
            except tk.TclError:
                pass
        finally:
            state["syncing"]["capture"] = False

    try:
        widget.after(0, _apply)
    except tk.TclError:
        pass


def sync_anchor_sliders() -> None:
    state = app_state.gui_control_state
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


def sync_overlay_sliders() -> None:
    state = app_state.gui_control_state
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
                widgets["offset_x"].set(int(app_state.info_overlay_offset["x"]))
                widgets["offset_y"].set(int(app_state.info_overlay_offset["y"]))
            except tk.TclError:
                pass
        finally:
            state["syncing"]["overlay"] = False

    try:
        widget.after(0, _apply)
    except tk.TclError:
        pass


# ---------- Overlay helpers ----------

def enforce_topmost(window: tk.Toplevel, interval_ms: int = 1500) -> None:
    """Continuously lift the overlay window so it stays above focused apps."""
    if not window.winfo_exists():
        return
    try:
        window.attributes("-topmost", True)  # type: ignore[attr-defined]
        window.lift()  # type: ignore[attr-defined]
    except tk.TclError:
        return
    window.after(interval_ms, lambda: enforce_topmost(window, interval_ms))


def create_overlay_window(width: int, height: int, left: int, top: int) -> tk.Toplevel:
    """Create a transparent always-on-top overlay window."""
    window = tk.Toplevel()
    window.attributes("-transparentcolor", "black")  # type: ignore[attr-defined]
    window.attributes("-topmost", True)  # type: ignore[attr-defined]
    window.overrideredirect(True)
    window.configure(bg="black")
    window.geometry(f"{width}x{height}+{left}+{top}")
    enforce_topmost(window)
    return window


def toggle_border() -> None:
    """Toggle visibility of the debug red border."""
    app_state.show_border = not app_state.show_border
    if app_state.border_canvas:
        app_state.border_canvas.itemconfig(
            "border", state="normal" if app_state.show_border else "hidden"
        )


def update_overlay_label(info: OverlayInfo, *, code: Optional[str] = None, raw_text: Optional[str] = None) -> None:
    """Update the floating label with the latest scan result."""
    message = ""
    if info:
        name = info.get("name", "")
        deposits = info.get("deposits")
        message = f"{name} x{deposits}" if deposits is not None else name

    app_state.overlay_text = message
    if message:
        app_state.last_overlay_time = time.time()
    else:
        app_state.last_overlay_time = 0
    if app_state.info_overlay_canvas and app_state.info_text_id:
        app_state.info_overlay_canvas.itemconfig(
            app_state.info_text_id, text=app_state.overlay_text, fill=app_state.label_color
        )
        logger.debug(f"Overlay updated: '{app_state.overlay_text}' (code={code}, raw_text={raw_text})")


def compute_info_overlay_geometry(screen_width: int, screen_height: int) -> Tuple[int, int, int, int]:
    overlay_width = max(400, min(800, screen_width - 40))
    overlay_height = 120
    base_left = max(0, (screen_width - overlay_width) // 2)
    base_top = max(0, int(screen_height * 0.35) - overlay_height // 2)

    offset_x = int(app_state.info_overlay_offset.get("x", 0))
    offset_y = int(app_state.info_overlay_offset.get("y", 0))

    max_left = max(0, screen_width - overlay_width)
    max_top = max(0, screen_height - overlay_height)

    left = min(max(0, base_left + offset_x), max_left)
    top = min(max(0, base_top + offset_y), max_top)
    return overlay_width, overlay_height, left, top


def reposition_info_overlay() -> None:
    if not app_state.info_overlay_root or not app_state.info_overlay_canvas or not app_state.info_text_id:
        return
    if not app_state.info_overlay_root.winfo_exists():
        return

    try:
        screen_width = app_state.info_overlay_root.winfo_screenwidth()
        screen_height = app_state.info_overlay_root.winfo_screenheight()
    except tk.TclError:
        geo = app_state.info_overlay_geometry
        screen_width = geo.get("screen_width") or 1920
        screen_height = geo.get("screen_height") or 1080

    overlay_width, overlay_height, left, top = compute_info_overlay_geometry(screen_width, screen_height)

    app_state.info_overlay_root.geometry(f"{overlay_width}x{overlay_height}+{left}+{top}")
    app_state.info_overlay_canvas.config(width=overlay_width, height=overlay_height)
    app_state.info_overlay_canvas.coords(app_state.info_text_id, overlay_width // 2, overlay_height // 2)
    app_state.info_overlay_canvas.itemconfig(app_state.info_text_id, width=overlay_width - 60)

    app_state.info_overlay_geometry.update({
        "screen_width": screen_width,
        "screen_height": screen_height,
        "width": overlay_width,
        "height": overlay_height,
    })


def start_label_timeout(window: Optional[tk.Toplevel]) -> None:
    """Background loop to clear overlay label if no update for 10s."""
    if app_state.info_overlay_canvas and app_state.info_text_id:
        if app_state.last_overlay_time and (time.time() - app_state.last_overlay_time > 10):
            app_state.info_overlay_canvas.itemconfig(app_state.info_text_id, text="")
            app_state.last_overlay_time = 0

    if window and window.winfo_exists():
        window.after(500, lambda: start_label_timeout(window))


def choose_label_color() -> None:
    from tkinter import colorchooser
    color = colorchooser.askcolor(title="Choose Label Color")[1]
    if color:
        app_state.label_color = color
        if app_state.info_overlay_canvas and app_state.info_text_id:
            app_state.info_overlay_canvas.itemconfig(app_state.info_text_id, fill=app_state.label_color)


# ---------- Capture overlay ----------

def _compute_capture_overlay_layout() -> Dict[str, int]:
    cap_w = int(app_state.cap_region["width"])
    cap_h = int(app_state.cap_region["height"])
    padding_x, padding_y = 100, 40

    overlay_width = cap_w + padding_x
    overlay_height = cap_h + padding_y
    left = int(app_state.cap_region["left"]) - (padding_x // 2)
    top = int(app_state.cap_region["top"]) - padding_y

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
    if not app_state.capture_overlay_canvas or not app_state.capture_rect_id or not app_state.capture_overlay_root:
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

    last = app_state.capture_overlay_last_layout
    size_changed = (
        force
        or last["overlay_width"] != overlay_width
        or last["overlay_height"] != overlay_height
    )
    pos_changed = force or last["left"] != left or last["top"] != top
    rect_changed = force or last["cap_w"] != cap_w or last["cap_h"] != cap_h

    if size_changed:
        app_state.capture_overlay_canvas.config(width=overlay_width, height=overlay_height)

    if rect_changed:
        app_state.capture_overlay_canvas.coords(
            app_state.capture_rect_id,
            padding_x // 2, padding_y,
            padding_x // 2 + cap_w, padding_y + cap_h,
        )

    if size_changed or pos_changed:
        app_state.capture_overlay_root.geometry(f"{overlay_width}x{overlay_height}+{left}+{top}")
        try:
            app_state.capture_overlay_root.lift()
        except tk.TclError:
            pass

    last.update({
        "overlay_width": overlay_width,
        "overlay_height": overlay_height,
        "left": left,
        "top": top,
        "cap_w": cap_w,
        "cap_h": cap_h,
    })


def _animate_capture_overlay() -> None:
    if (
        not app_state.capture_overlay_root
        or not app_state.capture_overlay_canvas
        or not app_state.capture_rect_id
        or not app_state.capture_overlay_root.winfo_exists()
    ):
        app_state.capture_overlay_animation_job = None
        return

    try:
        _apply_capture_overlay_layout()
    except tk.TclError:
        app_state.capture_overlay_animation_job = None
        return

    try:
        app_state.capture_overlay_animation_job = app_state.capture_overlay_root.after(
            33, _animate_capture_overlay
        )
    except tk.TclError:
        app_state.capture_overlay_animation_job = None


def start_capture_overlay_animation(*, force: bool = False) -> None:
    if not app_state.capture_overlay_root or not app_state.capture_overlay_canvas or not app_state.capture_rect_id:
        return

    _apply_capture_overlay_layout(force=force)

    if app_state.capture_overlay_animation_job is None:
        try:
            app_state.capture_overlay_animation_job = app_state.capture_overlay_root.after(
                33, _animate_capture_overlay
            )
        except tk.TclError:
            app_state.capture_overlay_animation_job = None


def stop_capture_overlay_animation() -> None:
    if app_state.capture_overlay_animation_job is not None and app_state.capture_overlay_root:
        try:
            app_state.capture_overlay_root.after_cancel(app_state.capture_overlay_animation_job)
        except (tk.TclError, ValueError):
            pass
    app_state.capture_overlay_animation_job = None

    app_state.capture_overlay_last_layout.update({
        "overlay_width": None, "overlay_height": None,
        "left": None, "top": None,
        "cap_w": None, "cap_h": None,
    })


def show_capture_overlay() -> None:
    if app_state.capture_overlay_root and app_state.capture_overlay_root.winfo_exists():
        try:
            stop_capture_overlay_animation()
            app_state.capture_overlay_root.destroy()
        except tk.TclError:
            pass
        app_state.capture_overlay_canvas = None
        app_state.capture_rect_id = None
        app_state.border_canvas = None

    cap_w, cap_h = int(app_state.cap_region["width"]), int(app_state.cap_region["height"])
    padding_x, padding_y = 100, 40

    overlay_width = cap_w + padding_x
    overlay_height = cap_h + padding_y
    left = int(app_state.cap_region["left"]) - (padding_x // 2)
    top = int(app_state.cap_region["top"]) - padding_y

    app_state.capture_overlay_root = create_overlay_window(overlay_width, overlay_height, left, top)

    app_state.capture_overlay_canvas = tk.Canvas(
        app_state.capture_overlay_root,
        width=overlay_width, height=overlay_height,
        bg="black", highlightthickness=0,
    )
    app_state.capture_overlay_canvas.pack()
    app_state.border_canvas = app_state.capture_overlay_canvas

    app_state.capture_rect_id = app_state.capture_overlay_canvas.create_rectangle(
        padding_x // 2, padding_y,
        padding_x // 2 + cap_w, padding_y + cap_h,
        outline="red", width=3, tags=("border",),
    )

    start_capture_overlay_animation(force=True)


def show_info_overlay(screen_width: int, screen_height: int) -> None:
    """Display a floating status overlay near the top-center of the screen."""
    if app_state.info_overlay_root and app_state.info_overlay_root.winfo_exists():
        try:
            app_state.info_overlay_root.destroy()
        except tk.TclError:
            pass
        app_state.info_overlay_canvas = None
        app_state.info_text_id = None

    overlay_width, overlay_height, left, top = compute_info_overlay_geometry(screen_width, screen_height)

    app_state.info_overlay_root = create_overlay_window(overlay_width, overlay_height, left, top)

    app_state.info_overlay_canvas = tk.Canvas(
        app_state.info_overlay_root,
        width=overlay_width, height=overlay_height,
        bg="black", highlightthickness=0,
    )
    app_state.info_overlay_canvas.pack()

    app_state.info_text_id = app_state.info_overlay_canvas.create_text(
        overlay_width // 2, overlay_height // 2,
        text=app_state.overlay_text,
        fill=app_state.label_color,
        font=("Arial", 18, "bold"),
        width=overlay_width - 60,
        justify="center",
    )

    app_state.info_overlay_geometry.update({
        "screen_width": screen_width,
        "screen_height": screen_height,
        "width": overlay_width,
        "height": overlay_height,
    })

    start_label_timeout(app_state.info_overlay_root)


def update_capture_overlay_region() -> None:
    start_capture_overlay_animation(force=True)


def show_anchor_overlay() -> None:
    if not app_state.anchor_overlay_visible:
        return

    if app_state.anchor_overlay_root and app_state.anchor_overlay_root.winfo_exists():
        try:
            app_state.anchor_overlay_root.destroy()
        except tk.TclError:
            pass
        app_state.anchor_overlay_canvas = None
        app_state.anchor_rect_id = None

    pad = 40
    width = int(app_state.anchor_region["width"]) + pad
    height = int(app_state.anchor_region["height"]) + pad
    left = int(app_state.anchor_region["left"]) - (pad // 2)
    top = int(app_state.anchor_region["top"]) - (pad // 2)

    app_state.anchor_overlay_root = create_overlay_window(width, height, left, top)

    app_state.anchor_overlay_canvas = tk.Canvas(
        app_state.anchor_overlay_root,
        width=width, height=height,
        bg="black", highlightthickness=0,
    )
    app_state.anchor_overlay_canvas.pack()

    app_state.anchor_rect_id = app_state.anchor_overlay_canvas.create_rectangle(
        pad // 2, pad // 2,
        pad // 2 + int(app_state.anchor_region["width"]),
        pad // 2 + int(app_state.anchor_region["height"]),
        outline="#00d4ff", width=2,
    )

    app_state.anchor_overlay_canvas.create_text(
        width // 2, 5,
        text="ANCHOR REGION",
        fill="#00d4ff",
        font=("Arial", 12, "bold"),
        anchor="n",
    )


def update_anchor_overlay_region() -> None:
    if (
        not app_state.anchor_overlay_visible
        or not app_state.anchor_overlay_root
        or not app_state.anchor_overlay_canvas
        or not app_state.anchor_rect_id
    ):
        return

    pad = 40
    width = int(app_state.anchor_region["width"]) + pad
    height = int(app_state.anchor_region["height"]) + pad
    left = int(app_state.anchor_region["left"]) - (pad // 2)
    top = int(app_state.anchor_region["top"]) - (pad // 2)

    app_state.anchor_overlay_canvas.config(width=width, height=height)
    app_state.anchor_overlay_canvas.coords(
        app_state.anchor_rect_id,
        pad // 2, pad // 2,
        pad // 2 + int(app_state.anchor_region["width"]),
        pad // 2 + int(app_state.anchor_region["height"]),
    )
    app_state.anchor_overlay_root.geometry(f"{width}x{height}+{left}+{top}")
    try:
        app_state.anchor_overlay_root.lift()
    except tk.TclError:
        pass


def hide_anchor_overlay() -> None:
    if app_state.anchor_overlay_root and app_state.anchor_overlay_root.winfo_exists():
        try:
            app_state.anchor_overlay_root.destroy()
        except tk.TclError:
            pass

    app_state.anchor_overlay_root = None
    app_state.anchor_overlay_canvas = None
    app_state.anchor_rect_id = None


def show_overlay(screen_width: Optional[int] = None, screen_height: Optional[int] = None) -> None:
    show_anchor_overlay()
    show_capture_overlay()

    if screen_width is None or screen_height is None:
        try:
            if app_state.capture_overlay_root and app_state.capture_overlay_root.winfo_exists():
                screen_width = app_state.capture_overlay_root.winfo_screenwidth()
                screen_height = app_state.capture_overlay_root.winfo_screenheight()
        except tk.TclError:
            screen_width = screen_width or 1920
            screen_height = screen_height or 1080

    if screen_width is not None and screen_height is not None:
        show_info_overlay(screen_width, screen_height)


def update_overlay_region() -> None:
    update_anchor_overlay_region()
    update_capture_overlay_region()
