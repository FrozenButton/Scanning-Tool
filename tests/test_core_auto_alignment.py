from scanning_tool.config.settings import AnchorSettings, CaptureSettings
from scanning_tool.core.auto_alignment import perform_auto_alignment
from scanning_tool.runtime.scan_state import AlignmentInfo


class DummyTracker:
    def __init__(self) -> None:
        self.threshold = 0.0
        self.locate_region = None

    def set_threshold(self, threshold: float) -> None:
        self.threshold = threshold

    def locate_anchor(self, region):
        self.locate_region = region
        return {
            "match_left": 100.0,
            "match_top": 200.0,
            "score": 0.95,
            "template": "dummy.png",
            "template_width": 50.0,
            "template_height": 50.0,
        }


def test_perform_auto_alignment_updates_capture_region_and_callbacks():
    tracker = DummyTracker()
    anchor_settings = AnchorSettings(
        anchor_region={"left": 1, "top": 2, "width": 20, "height": 30},
        anchor_offset={"x": 10, "y": 15},
        anchor_threshold=0.8,
        auto_align_enabled=True,
    )
    capture_settings = CaptureSettings(cap_region={"left": 0, "top": 0, "width": 100, "height": 80})
    last_alignment_info = AlignmentInfo()

    sync_calls = []
    update_calls = []

    def sync_capture_sliders() -> None:
        sync_calls.append("sync")

    def update_capture_overlay_region() -> None:
        update_calls.append("update")

    result = perform_auto_alignment(
        tracker,
        anchor_settings,
        capture_settings,
        last_alignment_info,
        sync_capture_sliders,
        update_capture_overlay_region,
    )

    assert result is True
    assert last_alignment_info.matched is True
    assert last_alignment_info.template == "dummy.png"
    assert last_alignment_info.score == 0.95
    assert sync_calls == ["sync"]
    assert update_calls == ["update"]
    assert capture_settings.cap_region["left"] == 85
    assert capture_settings.cap_region["top"] == 215


def test_perform_auto_alignment_returns_false_when_disabled():
    tracker = DummyTracker()
    anchor_settings = AnchorSettings(auto_align_enabled=False)
    capture_settings = CaptureSettings()
    last_alignment_info = AlignmentInfo()

    result = perform_auto_alignment(
        tracker,
        anchor_settings,
        capture_settings,
        last_alignment_info,
        lambda: None,
        lambda: None,
    )

    assert result is False
    assert last_alignment_info.matched is False
