"""Glass theme styling for the Tkinter GUI."""

from typing import Dict

import tkinter as tk
from tkinter import ttk

from PIL import Image, ImageDraw, ImageTk


GlassPalette = Dict[str, str]


_PALETTE: GlassPalette = {
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


def apply_glass_theme(root: tk.Tk) -> GlassPalette:
    """Apply a holographic glass-inspired theme to the Tkinter UI."""
    colors = dict(_PALETTE)

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

    slider_normal = _make_slider_image(colors["knob"], colors["knob_outline"])
    slider_active = _make_slider_image(colors["knob_active"], colors["accent"])
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


def _make_slider_image(fill: str, outline: str) -> ImageTk.PhotoImage:
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


def style_spinbox(spinbox: tk.Spinbox, colors: GlassPalette) -> None:
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
