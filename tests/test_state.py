from scanning_tool.state import app_state


def test_app_state_has_expected_components():
    assert app_state.settings is not None
    assert app_state.scan_state.continuous_mode is False
    assert app_state.overlay_state is not None
    assert app_state.control_state is not None
    assert app_state.service_state is not None
