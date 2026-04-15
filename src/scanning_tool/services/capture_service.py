"""Service for handling screen capture and OCR processing."""

import re
import time
from threading import Thread

import mss
from PIL import Image

from scanning_tool.state import app_state
from scanning_tool.ocr import ocr_with_ollama
from scanning_tool.deposits import extract_code_from_text, lookup_deposit
from scanning_tool.core.auto_alignment import perform_auto_alignment
from scanning_tool.gui.overlays import update_capture_overlay_region, update_overlay_label, sync_capture_sliders
from scanning_tool.runtime.scan_state import LastResult

from scanning_tool.interfaces.services import ICaptureService


class CaptureService(ICaptureService):
    """Service for capturing screen regions and processing OCR results."""
    
    def __init__(self):
        self._status_callback = None
    
    def capture_once(self, status_callback=None) -> None:
        """Capture one scan from the capture region and update overlay."""
        self._status_callback = status_callback
        self._do_capture()
    
    def _highlight_numbers(self, text: str) -> str:
        """Wrap numbers in <yellow> tags for loguru."""
        return re.sub(r"(\d+)", r"<yellow>\1</yellow>", text)
    
    def _do_capture(self):
        """Internal capture logic."""
        if self._status_callback:
            self._status_callback("Aligning region...")
        
        auto_aligned = perform_auto_alignment(
            app_state.scan_state.anchor_tracker,
            app_state.settings.anchor,
            app_state.settings.capture,
            app_state.scan_state.last_alignment_info,
            sync_capture_sliders,
            update_capture_overlay_region,
        )
        
        if app_state.settings.anchor.auto_align_enabled:
            from loguru import logger
            logger.debug("Auto alignment %s before capture.", "succeeded" if auto_aligned else "did not match")
        
        cap_region = app_state.settings.capture.cap_region
        with mss.mss() as sct:
            monitor = {
                "left": cap_region["left"],
                "top": cap_region["top"],
                "width": cap_region["width"],
                "height": cap_region["height"],
            }
            img = sct.grab(monitor)
            pil_img = Image.frombytes("RGB", img.size, img.rgb)
        
        if self._status_callback:
            self._status_callback("Loading OCR model (may take a moment)...")
        
        from loguru import logger
        logger.debug("Loading OCR model for scan...")
        
        try:
            raw_text = ocr_with_ollama(pil_img)
            code, raw = extract_code_from_text(raw_text)
            info = lookup_deposit(code)
            app_state.scan_state.last_result = LastResult(code=code, code_raw=raw, info=info, raw_text=raw_text)
            update_overlay_label(info, code=code, raw_text=raw or raw_text)
            
            result = app_state.scan_state.last_result
            logger.info(
                f"Scan result: LastResult("
                f"code={result.code}, "
                f"code_raw={result.code_raw}, "
                f"info={result.info}, "
                f"confidence={result.confidence}, "
                f"raw_text={result.raw_text}"
                ")"
            )
            
            if self._status_callback:
                self._status_callback("Scan complete.")
                
        except Exception as e:
            logger.error(f"OCR/model error: {e}")
            if self._status_callback:
                self._status_callback(f"OCR/model error: {e}")
    
    def toggle_continuous(self) -> None:
        """Toggle continuous scanning mode."""
        from scanning_tool.state import app_state
        app_state.scan_state.continuous_mode = not app_state.scan_state.continuous_mode
        from loguru import logger
        logger.info(f"Continuous mode: {app_state.scan_state.continuous_mode}")
        
        if app_state.scan_state.continuous_mode:
            Thread(target=self._continuous_scan_loop, daemon=True).start()
    
    def _continuous_scan_loop(self) -> None:
        """Run scans repeatedly until continuous_mode is turned off."""
        while True:
            from scanning_tool.state import app_state
            if not app_state.scan_state.continuous_mode:
                break
            self.capture_once()
            interval = max(0.1, float(app_state.settings.capture.continuous_capture_interval))
            time.sleep(interval)


# Create a singleton instance for backward compatibility
capture_service = CaptureService()