from scanning_tool.state import app_state
from scanning_tool.state_context import app_state as legacy_app_state


def test_app_state_is_shared_between_state_packages():
    assert app_state is legacy_app_state
    assert app_state.settings is not None
    assert app_state.scan_state.continuous_mode is False
    assert app_state.overlay_state is not None
    assert app_state.control_state is not None
    assert app_state.service_state is not None
