import json
from pathlib import Path

from scanning_tool.config.loader import load_config, save_config
from scanning_tool.state.context import create_app_context


def test_config_loader_roundtrips(tmp_path: Path) -> None:
    app_context = create_app_context()
    app_context.settings.capture.cap_region = {"left": 1, "top": 2, "width": 3, "height": 4}
    app_context.settings.overlay.label_color = "blue"
    app_context.settings.anchor.anchor_threshold = 0.5

    config_path = tmp_path / "config.json"
    save_config(app_context, config_path)

    loaded_context = create_app_context()
    load_config(loaded_context, config_path)

    assert loaded_context.settings.capture.cap_region == {
        "left": 1,
        "top": 2,
        "width": 3,
        "height": 4,
    }
    assert loaded_context.settings.overlay.label_color == "blue"
    assert loaded_context.settings.anchor.anchor_threshold == 0.5
