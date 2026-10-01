"""Check the CLI evaluation transport without Excel or an external model."""

from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp.types import TextContent


class CliEvaluationTransportTests(unittest.IsolatedAsyncioTestCase):
    async def test_wrapper_initializes_and_executes_with_locked_mcp_sdk(self):
        server = StdioServerParameters(
            command=sys.executable,
            args=[
                str(Path(__file__).with_name("cli_mcp_server.py")),
                "--command", sys.executable, "--tool-prefix", "test", "--shell", "none",
            ],
        )
        async with stdio_client(server) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                tools = await session.list_tools()
                self.assertEqual([tool.name for tool in tools.tools], ["test_execute"])
                called = await session.call_tool("test_execute", {"args": '-c "print(42)"'})
                self.assertFalse(called.is_error)
                content = called.content[0]
                assert isinstance(content, TextContent), content
                result = json.loads(content.text)
                self.assertEqual(result["exit_code"], 0, result)
                self.assertEqual(result["stdout"].strip(), "42")


if __name__ == "__main__":
    unittest.main()
