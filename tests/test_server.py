import pytest

from msproject_mcp.com_bridge import ComBridge
from msproject_mcp.server import create_server
from fake_project import FakeApp


@pytest.fixture
def server():
    app = FakeApp()
    return create_server(ComBridge(app_factory=lambda visible: app))


async def test_tools_registered(server):
    names = {t.name for t in await server.list_tools()}
    assert {"get_status", "add_task", "link_tasks", "assign_resource", "save_project"} <= names


async def test_tool_round_trip(server):
    await server.call_tool("create_project", {"title": "Demo"})
    await server.call_tool("add_task", {"name": "Kickoff", "duration": "0d"})
    _, result = await server.call_tool("list_tasks", {})
    assert result["tasks"][0]["name"] == "Kickoff"


async def test_tool_errors_are_reported(server):
    from mcp.server.fastmcp.exceptions import ToolError

    with pytest.raises(ToolError, match="No project is open"):
        await server.call_tool("list_tasks", {})
