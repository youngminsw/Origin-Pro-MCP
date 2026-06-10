import pytest

pytest.importorskip("mcp")


def test_mcp_server_uses_quiet_stdio_log_level() -> None:
    from origin_pro_mcp.app import mcp

    assert mcp.settings.log_level == "ERROR"
