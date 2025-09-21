import time
import re
import json
import io
import base64
import os
import sys
from threading import Thread
from PIL import Image
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

def ensure_ollama_installed():
    """
    Check if Ollama is installed on the system (cross-platform).
    If not found, direct user to download from official website rather than using embedded installer.
    Works on Windows and Linux.
    """
    if not shutil.which("ollama"):
        print("Ollama not found on your system.")
        print("Please install Ollama from the official website: https://ollama.com/")
        
        # Show a popup with installation instructions
        messagebox.showinfo(
            "Ollama Required",
            "Ollama is required but not installed.\n\n"
            "Please visit https://ollama.com/ to download and install Ollama for your operating system.\n\n"
            "After installation, restart this program to continue."
        )
        
        # Open the website in the default browser
        try:
            webbrowser.open("https://ollama.com/")
        except Exception as e:
            print(f"Could not open browser automatically: {e}")
            print("Please manually visit https://ollama.com/ to download Ollama")
        
        # Wait for user acknowledgment
        input("\nPress ENTER to close after installing Ollama...")
        sys.exit(0)

    else:
        try:
            version = subprocess.check_output(["ollama", "--version"], text=True).strip()
            print(f"Ollama found: {version}")
        except Exception as e:
            print("Error checking Ollama:", e)
            sys.exit("Please install Ollama and rerun this program.")

def ensure_model_installed(model="qwen2.5vl:3b"):
    """Ensure the Ollama model is pulled locally."""
    try:
        result = subprocess.run(["ollama", "list"], capture_output=True, text=True)
        if model not in result.stdout:
            print(f"Model {model} not found. Pulling now...")
            subprocess.run(["ollama", "pull", model], check=True)
            print(f"Model {model} installed successfully.")
        else:
            print(f"Model {model} already installed.")
    except Exception as e:
        print("Error ensuring model:", e)
        sys.exit("Failed to ensure Ollama model.")



# ---------- CONFIG ----------
CONFIG_FILE = "config.json"

CAP_REGION = {"left": 1260, "top": 310, "width": 160, "height": 30}
label_color = "yellow"
MIN_CONFIDENCE = 0.65
DEBUG_SHOW_OVERLAY = True
OLLAMA_MODEL = "qwen2.5vl:3b"   # vision model

# Regex for codes
CODE_RE = re.compile(
    r"(?:[A-Za-z]?-?\d[\d,\.]{1,10}|\d{2,10})",
    re.IGNORECASE
)

last_result = {"code": None, "code_raw": None, "info": None,
               "confidence": 0.0, "raw_text": ""}


# ---------- Config Handling ----------
def load_config():
    global CAP_REGION, label_color
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                data = json.load(f)
                CAP_REGION = data.get("CAP_REGION", CAP_REGION)
                label_color = data.get("label_color", label_color)
        except (json.JSONDecodeError, OSError) as e:
            print(f"Config file invalid or empty, resetting: {e}")
            save_config()
    else:
        save_config()



def save_config():
    global CAP_REGION, label_color
    data = {"CAP_REGION": CAP_REGION, "label_color": label_color}
    with open(CONFIG_FILE, "w") as f:
        json.dump(data, f, indent=4)
    print("Config saved.")


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
    2000: {"key": "SALVAGE", "display_name": "Metal Pannals", "rarity": "common", "category": "Savlage"},
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
    try:
        response = ollama.chat(
            model=model,
            messages=[{
                "role": "user",
                "content": "Extract the numeric code shown in this image. Only return the code, no extra words.",
                "images": [img_bytes],
            }],
        )
        return response["message"]["content"].strip()
    except Exception as e:
        print("Ollama OCR error:", e)
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
overlay_canvas = None
overlay_text_id = None
overlay_text = ""
last_overlay_time = 0
root_overlay = None


def toggle_border():
    """Toggle visibility of the debug red border."""
    global show_border, border_canvas
    show_border = not show_border
    if border_canvas:
        border_canvas.itemconfig("border", state="normal" if show_border else "hidden")


def update_overlay_label(info):
    """Update the overlay label with deposit info and reset timeout timer."""
    global overlay_text, overlay_text_id, overlay_canvas, last_overlay_time
    if info:
        overlay_text = f"{info['name']} x{info['deposits']}" if "deposits" in info else info["name"]
        last_overlay_time = time.time()
        if overlay_canvas and overlay_text_id:
            overlay_canvas.itemconfig(overlay_text_id, text=overlay_text, fill=label_color)


def start_label_timeout(root):
    """Background loop to clear overlay label if no update for 10s."""
    global overlay_text_id, overlay_canvas, last_overlay_time
    if overlay_canvas and overlay_text_id:
        if last_overlay_time and (time.time() - last_overlay_time > 10):
            overlay_canvas.itemconfig(overlay_text_id, text="")
            last_overlay_time = 0
    root.after(500, lambda: start_label_timeout(root))



# ---------- GUI + Overlay ----------
def choose_label_color():
    global label_color, overlay_canvas, overlay_text_id
    color = colorchooser.askcolor(title="Choose Label Color")[1]
    if color:
        label_color = color
        if overlay_canvas and overlay_text_id:
            overlay_canvas.itemconfig(overlay_text_id, fill=label_color)


def show_overlay():
    global border_canvas, overlay_canvas, overlay_text_id, rect_id, root_overlay

    cap_w, cap_h = int(CAP_REGION['width']), int(CAP_REGION['height'])
    padding_x, padding_y = 100, 40

    overlay_width = cap_w + padding_x
    overlay_height = cap_h + padding_y
    left = int(CAP_REGION['left']) - (padding_x // 2)
    top = int(CAP_REGION['top']) - padding_y

    root_overlay = tk.Toplevel()
    root_overlay.attributes("-transparentcolor", "black")
    root_overlay.attributes("-topmost", True)
    root_overlay.overrideredirect(True)
    root_overlay.configure(bg="black")
    root_overlay.geometry(f"{overlay_width}x{overlay_height}+{left}+{top}")

    overlay_canvas = tk.Canvas(root_overlay, width=overlay_width, height=overlay_height,
                               bg="black", highlightthickness=0)
    overlay_canvas.pack()
    border_canvas = overlay_canvas

    rect_id = overlay_canvas.create_rectangle(
        padding_x // 2, padding_y,
        padding_x // 2 + cap_w, padding_y + cap_h,
        outline="red", width=3, tags=("border",)
    )

    overlay_text_id = overlay_canvas.create_text(
        overlay_width // 2, 5,
        text="", fill=label_color, font=("Arial", 14, "bold"),
        width=overlay_width - 20, anchor="n"
    )
    start_label_timeout(root_overlay)



def update_overlay_region():
    global overlay_canvas, rect_id, root_overlay
    if not overlay_canvas or not rect_id:
        return
    cap_w, cap_h = int(CAP_REGION['width']), int(CAP_REGION['height'])
    padding_x, padding_y = 100, 40
    overlay_width = cap_w + padding_x
    overlay_height = cap_h + padding_y
    left = int(CAP_REGION['left']) - (padding_x // 2)
    top = int(CAP_REGION['top']) - padding_y

    overlay_canvas.coords(rect_id, padding_x // 2, padding_y,
                          padding_x // 2 + cap_w, padding_y + cap_h)
    root_overlay.geometry(f"{overlay_width}x{overlay_height}+{left}+{top}")


def launch_gui():
    def update_region_from_sliders(*args):
        CAP_REGION["left"] = int(slider_left.get())
        CAP_REGION["top"] = int(slider_top.get())
        CAP_REGION["width"] = int(slider_width.get())
        CAP_REGION["height"] = int(slider_height.get())
        lbl_status.config(text=f"CAP_REGION = {CAP_REGION}")

    def on_close():
        save_config()
        try:
            if root_overlay:
                root_overlay.destroy()
        except:
            pass
        root.destroy()

    root = tk.Tk()
    root.title("Star Citizen Scanner Control")
    root.protocol("WM_DELETE_WINDOW", on_close)

    frm_region = ttk.LabelFrame(root, text="Capture Region")
    frm_region.pack(fill="x", padx=5, pady=5)

    slider_left = tk.Scale(frm_region, from_=0, to=3000, orient="horizontal",
                           label="Left", command=update_region_from_sliders)
    slider_left.set(CAP_REGION["left"])
    slider_left.pack(fill="x")

    slider_top = tk.Scale(frm_region, from_=0, to=2000, orient="horizontal",
                          label="Top", command=update_region_from_sliders)
    slider_top.set(CAP_REGION["top"])
    slider_top.pack(fill="x")

    slider_width = tk.Scale(frm_region, from_=50, to=1000, orient="horizontal",
                            label="Width", command=update_region_from_sliders)
    slider_width.set(CAP_REGION["width"])
    slider_width.pack(fill="x")

    slider_height = tk.Scale(frm_region, from_=20, to=500, orient="horizontal",
                             label="Height", command=update_region_from_sliders)
    slider_height.set(CAP_REGION["height"])
    slider_height.pack(fill="x")

    frm_ctrl = ttk.LabelFrame(root, text="Controls")
    frm_ctrl.pack(fill="x", padx=5, pady=5)

    tk.Button(frm_ctrl, text="Single Scan", command=capture_once).pack(side="left", padx=5)
    tk.Button(frm_ctrl, text="Loop Toggle", command=toggle_continuous).pack(side="left", padx=5)
    tk.Button(frm_ctrl, text="Update Overlay", command=update_overlay_region).pack(side="left", padx=5)
    tk.Button(frm_ctrl, text="Set Label Color", command=choose_label_color).pack(side="left", padx=5)
    tk.Button(frm_ctrl, text="Save Config", command=save_config).pack(side="left", padx=5)
    tk.Button(frm_ctrl, text="Toggle Border", command=toggle_border).pack(side="left", padx=5)


    lbl_status = tk.Label(root, text="Ready.", anchor="w", justify="left")
    lbl_status.pack(fill="x", padx=5, pady=5)

    show_overlay()
    root.mainloop()


# ---------- Scanning Functions ----------
def capture_once():
    """Capture one scan from CAP_REGION and update overlay."""
    global last_result
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
    update_overlay_label(info)
    print("Scan result:", last_result)


def toggle_continuous():
    """Toggle continuous scanning mode."""
    global continuous_mode
    continuous_mode = not continuous_mode
    print("Continuous mode:", continuous_mode)
    if continuous_mode:
        Thread(target=continuous_scan_loop, daemon=True).start()


def continuous_scan_loop():
    """Run scans repeatedly until continuous_mode is turned off."""
    while continuous_mode:
        capture_once()
        time.sleep(2)  # adjust scan interval



# ---------- Flask / Hotkeys ----------
template_folder = resource_path("templates")
app = Flask(__name__, template_folder=template_folder)

@app.route("/")
def index():
    return render_template("overlay.html")



@app.route("/status")
def status():
    return jsonify({"region": CAP_REGION, "label_color": label_color, "last": last_result})


def hotkey_listener():
    """Set up hotkey listeners with cross-platform error handling."""
    try:
        keyboard.add_hotkey("7", capture_once)
        keyboard.add_hotkey("ctrl+7", toggle_continuous)
        keyboard.add_hotkey("8", toggle_border)
        print("Hotkeys registered: '7' for single scan, 'Ctrl+7' for continuous toggle, '8' for border toggle")
        keyboard.wait()
    except Exception as e:
        print(f"Could not set up global hotkeys: {e}")
        print("Note: Linux Support is being tested.")


# ---------- Main ----------
if __name__ == "__main__":
    # Ensure Ollama + model before starting
    ensure_ollama_installed()
    ensure_model_installed("qwen2.5vl:3b")

    load_config()
    Thread(target=hotkey_listener, daemon=True).start()
    Thread(target=lambda: app.run(host="127.0.0.1", port=5000, debug=False), daemon=True).start()
    launch_gui()
