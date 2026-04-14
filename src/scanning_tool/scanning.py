"""Screen capture and scanning logic."""

import logging
import time
from threading import Thread

import mss
from PIL import Image

from scanning_tool.state import app_state
from scanning_tool.ocr import ocr_with_ollama
from scanning_tool.deposits import extract_code_from_text, lookup_deposit
from scanning_tool.anchor import perform_auto_alignment
from scanning_tool.overlay import update_overlay_label

logger = logging.getLogger("scanning_tool")


def capture_once() -> None:
    """Capture one scan from the capture region and update overlay."""
    auto_aligned = perform_auto_alignment()
    if app_state.auto_align_enabled:
        logger.debug("Auto alignment %s before capture.", "succeeded" if auto_aligned else "did not match")
    with mss.mss() as sct:
        monitor = {
            "left": app_state.cap_region["left"],
            "top": app_state.cap_region["top"],
            "width": app_state.cap_region["width"],
            "height": app_state.cap_region["height"],
        }
        img = sct.grab(monitor)
        pil_img = Image.frombytes("RGB", img.size, img.rgb)

    raw_text = ocr_with_ollama(pil_img)
    code, raw = extract_code_from_text(raw_text)
    info = lookup_deposit(code)

    app_state.last_result = {"code": code, "code_raw": raw, "info": info, "raw_text": raw_text}
    update_overlay_label(info, code=code, raw_text=raw or raw_text)
    logger.info(f"Scan result: {app_state.last_result}")


def toggle_continuous() -> None:
    """Toggle continuous scanning mode."""
    app_state.continuous_mode = not app_state.continuous_mode
    logger.info(f"Continuous mode: {app_state.continuous_mode}")
    if app_state.continuous_mode:
        Thread(target=continuous_scan_loop, daemon=True).start()


def continuous_scan_loop() -> None:
    """Run scans repeatedly until continuous_mode is turned off."""
    while app_state.continuous_mode:
        capture_once()
        interval = max(0.1, float(app_state.continuous_capture_interval))
        time.sleep(interval)
