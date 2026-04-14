import os
from pathlib import Path

from scanning_tool.config import ensure_anchor_directory, resource_path


def test_ensure_anchor_directory_creates_directory(tmp_path):
    target = tmp_path / "anchors"
    assert not target.exists()

    ensure_anchor_directory(str(target))

    assert target.exists()
    assert target.is_dir()


def test_resource_path_returns_valid_path_for_config():
    path = resource_path("config.json")
    assert isinstance(path, str)
    assert Path(path).name == "config.json"
    assert Path(path).suffix == ".json"
