"""Anchor template loading and matching."""

import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import cv2
import mss
import numpy as np

from scanning_tool.config import ensure_anchor_directory

logger = logging.getLogger(__name__)


class AnchorRegionTracker:
    """Manage template loading and anchor matching for auto alignment."""

    def __init__(self, template_dir: str, threshold: float = 0.82) -> None:
        self.template_dir = template_dir
        self.threshold = threshold
        self.templates: List[Tuple[str, np.ndarray]] = []
        self.last_loaded_count = 0
        self.load_templates()

    def set_threshold(self, threshold: float) -> None:
        self.threshold = threshold

    def set_directory(self, template_dir: str) -> int:
        self.template_dir = template_dir
        return self.load_templates()

    def load_templates(self) -> int:
        ensure_anchor_directory(self.template_dir)
        loaded: List[Tuple[str, np.ndarray]] = []
        directory = Path(self.template_dir)
        if not directory.exists():
            logger.debug(f"Anchor template directory does not exist: {directory}")
            self.templates = []
            self.last_loaded_count = 0
            return 0

        supported_ext = {".png", ".jpg", ".jpeg", ".bmp"}
        for path in sorted(directory.glob("**/*")):
            if path.suffix.lower() not in supported_ext or not path.is_file():
                continue
            image = cv2.imread(str(path), cv2.IMREAD_GRAYSCALE)
            if image is None:
                logger.warning(f"Failed to load anchor template: {path}")
                continue
            loaded.append((path.name, image))

        self.templates = loaded
        self.last_loaded_count = len(loaded)
        if self.last_loaded_count == 0:
            logger.warning(
                "No anchor templates were loaded. Head sway compensation will remain disabled until templates are added."
            )
        else:
            logger.info(f"Loaded {self.last_loaded_count} anchor templates from {directory}")
        return self.last_loaded_count

    def locate_anchor(self, region: Dict[str, int]) -> Optional[Dict[str, float]]:
        if not self.templates:
            return None

        monitor = {
            "left": int(region["left"]),
            "top": int(region["top"]),
            "width": int(region["width"]),
            "height": int(region["height"]),
        }

        with mss.mss() as sct:
            try:
                screenshot = sct.grab(monitor)
            except Exception as exc:
                logger.error(f"Anchor capture failed: {exc}")
                return None

        anchor_image = np.array(screenshot)
        if anchor_image.ndim == 3 and anchor_image.shape[2] == 4:
            anchor_gray = cv2.cvtColor(anchor_image, cv2.COLOR_BGRA2GRAY)
        else:
            anchor_gray = cv2.cvtColor(anchor_image, cv2.COLOR_BGR2GRAY)

        best_score = -1.0
        best_loc: Optional[Tuple[int, int]] = None
        best_template: Optional[Tuple[str, np.ndarray]] = None

        for template_name, template_img in self.templates:
            if anchor_gray.shape[0] < template_img.shape[0] or anchor_gray.shape[1] < template_img.shape[1]:
                continue
            res = cv2.matchTemplate(anchor_gray, template_img, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(res)
            if max_val > best_score:
                best_score = float(max_val)
                best_loc = max_loc
                best_template = (template_name, template_img)

        if best_loc is None or best_template is None:
            return None

        if best_score < self.threshold:
            logger.debug(
                f"Anchor match below threshold ({best_score:.3f} < {self.threshold:.3f}) using template {best_template[0]}"
            )
            return None

        match_left = monitor["left"] + best_loc[0]
        match_top = monitor["top"] + best_loc[1]
        return {
            "match_left": float(match_left),
            "match_top": float(match_top),
            "score": best_score,
            "template": best_template[0],
            "template_width": float(best_template[1].shape[1]),
            "template_height": float(best_template[1].shape[0]),
        }
