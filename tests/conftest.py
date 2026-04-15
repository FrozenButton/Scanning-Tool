"""Pytest configuration and fixture registry."""

import pytest
from unittest.mock import MagicMock

@pytest.fixture
def mock_logger():
    """Provides a mocked logger for testing."""
    return MagicMock()

@pytest.fixture
def sample_config_dict():
    """Provides a baseline config dictionary for testing logic."""
    return {
        "capture_region": {
            "left": 0, "top": 0, "width": 100, "height": 100
        }
    }
