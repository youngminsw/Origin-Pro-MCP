import asyncio
import os
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.requires_origin


def test_mcp_stdio_lists_tools_and_guards_labtalk() -> None:
    async def scenario() -> None:
        from mcp.client.session import ClientSession
        from mcp.client.stdio import StdioServerParameters, stdio_client

        env = os.environ.copy()
        env["PYTHONPATH"] = str(Path.cwd() / "src")
        params = StdioServerParameters(
            command=sys.executable,
            args=["-m", "origin_pro_mcp.server"],
            env=env,
            cwd=Path.cwd(),
        )
        async with stdio_client(params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                tools = await session.list_tools()
                tool_names = {tool.name for tool in tools.tools}

                assert "run_labtalk" in tool_names
                assert "create_graph" in tool_names
                assert "export_graph" in tool_names

                run_result = await session.call_tool(
                    "run_labtalk",
                    {"script": "double __mcp_stdio_test = 789;"},
                )
                assert run_result.content[0].text == (
                    "Executed successfully: double __mcp_stdio_test = 789;"
                )

                read_result = await session.call_tool(
                    "get_labtalk_variable",
                    {"name": "__mcp_stdio_test"},
                )
                assert read_result.content[0].text == "789.0"

                blocked_result = await session.call_tool(
                    "run_labtalk",
                    {"script": "doc -s;"},
                )
                assert "blocked by the safety guard" in blocked_result.content[0].text

    asyncio.run(scenario())
