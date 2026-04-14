import numpy as np
import cv2

from scanning_tool.core.anchor import AnchorRegionTracker


def test_load_templates_reads_supported_images(tmp_path):
    image_path = tmp_path / "template.png"
    image = np.zeros((16, 16, 3), dtype=np.uint8)
    cv2.imwrite(str(image_path), image)

    tracker = AnchorRegionTracker(str(tmp_path), threshold=0.5)

    assert tracker.last_loaded_count == 1
    assert tracker.templates[0][0] == "template.png"


def test_load_templates_returns_zero_for_empty_directory(tmp_path):
    tracker = AnchorRegionTracker(str(tmp_path), threshold=0.5)

    assert tracker.last_loaded_count == 0
    assert tracker.templates == []
