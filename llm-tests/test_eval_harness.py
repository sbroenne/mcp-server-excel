"""Offline checks of evaluation configuration and completion evidence."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import AsyncMock, patch
from types import SimpleNamespace

from copilot import CopilotSession
from pytest_skill_engineering.copilot.result import CopilotResult, ToolCall, Turn

from conftest import build_excel_cli_eval, build_excel_mcp_eval, assert_cli_exit_codes
from consent_scenarios import isolated_cli_servers
from skill_discovery import discover_skills


class DiscoveryPreflightTests(unittest.IsolatedAsyncioTestCase):
    async def test_required_native_tools_are_checked_without_model_messages(self):
        with tempfile.TemporaryDirectory() as directory:
            agent = build_excel_cli_eval("preflight", servers={}, working_directory=directory)
            for names in (["powershell", "skill"], ["skill"]):
                with self.subTest(names=names):
                    client = AsyncMock()
                    client.create_session.return_value.rpc.skills.list.return_value = SimpleNamespace(skills=[])
                    client.rpc.tools.list.return_value = SimpleNamespace(
                        tools=[SimpleNamespace(name=name) for name in names],
                    )
                    with patch("skill_discovery.CopilotClient", return_value=client):
                        if "powershell" in names:
                            self.assertEqual(await discover_skills(agent, required_tools=("powershell", "skill")), [])
                        else:
                            with self.assertRaisesRegex(AssertionError, "Missing native tools"):
                                await discover_skills(agent, required_tools=("powershell", "skill"))
                    client.stop.assert_awaited_once()

    async def test_package_setup_failure_blocks_discovery_before_a_model_send(self):
        with tempfile.TemporaryDirectory() as directory:
            agent = build_excel_cli_eval("preflight", servers={}, working_directory=directory, skill_dir=directory)
            with (
                patch("skill_discovery.load_skill", side_effect=ValueError("Invalid references entry")),
                patch("skill_discovery.CopilotClient", return_value=AsyncMock()) as client,
            ):
                with self.assertRaisesRegex(ValueError, "Invalid references entry"):
                    await discover_skills(agent)
                client.assert_not_called()

    async def test_accidental_model_send_during_discovery_is_rejected(self):
        async def create_session(**config):
            await CopilotSession.send(None, "must-not-send")

        with tempfile.TemporaryDirectory() as directory:
            agent = build_excel_cli_eval("preflight", servers={}, working_directory=directory)
            client = AsyncMock()
            client.create_session.side_effect = create_session
            with patch("skill_discovery.CopilotClient", return_value=client):
                with self.assertRaises(AssertionError):
                    await discover_skills(agent)
            client.stop.assert_awaited_once()

    async def test_preflight_loads_actual_directory_without_starting_excel_servers(self):
        with tempfile.TemporaryDirectory() as directory:
            servers = {"unused": {"command": "must-not-run", "args": []}}
            agent = build_excel_cli_eval("preflight", servers=servers, working_directory=directory, skill_dir=directory)
            client = AsyncMock()
            client.create_session.return_value.rpc.skills.list.return_value = SimpleNamespace(skills=[])
            with (
                patch("skill_discovery.load_skill") as load,
                patch("skill_discovery.CopilotClient", return_value=client),
            ):
                self.assertEqual(await discover_skills(agent), [])
            self.assertEqual(agent.mcp_servers, servers)
            load.assert_called_once_with(directory)
            self.assertEqual(client.create_session.call_args.kwargs["mcp_servers"], {})
            client.stop.assert_awaited_once()


class EvalHarnessTests(unittest.TestCase):
    def test_builders_use_current_single_attempt_api(self):
        for builder in (build_excel_cli_eval, build_excel_mcp_eval):
            with self.subTest(builder=builder.__name__):
                agent = builder("offline", servers={})
                self.assertEqual(agent.max_tool_calls, 80)
                self.assertEqual(agent.client_mode, "empty")
                self.assertFalse(agent.build_session_config()["enable_config_discovery"])
                self.assertFalse(agent.build_session_config()["enable_on_demand_instruction_discovery"])

    def test_agent_workspace_is_not_the_source_repository(self):
        with tempfile.TemporaryDirectory() as directory:
            agent = build_excel_cli_eval("offline", servers={}, working_directory=directory)
            self.assertEqual(agent.working_directory, str(Path(directory).resolve()))

    def test_private_cli_uses_the_same_task_directory_as_the_agent(self):
        servers = {"excel-cli": {"args": ["wrapper.py", "--cwd", "old"], "env": {}}}
        isolated = isolated_cli_servers(servers, "private-pipe", working_directory="task")
        self.assertEqual(isolated["excel-cli"]["args"], ["wrapper.py", "--cwd", "task"])
        self.assertEqual(isolated["excel-cli"]["cwd"], "task")
        self.assertEqual(servers["excel-cli"]["args"], ["wrapper.py", "--cwd", "old"])

    def test_explicit_skills_are_enabled_without_ambient_discovery(self):
        with tempfile.TemporaryDirectory() as directory:
            skill = Path(directory) / "explicit-skill"
            skill.mkdir()
            (skill / "SKILL.md").write_text(
                "---\nname: explicit-skill\ndescription: Offline fixture guidance.\n---\nUse the task tools.\n",
                encoding="utf-8",
            )
            for builder in (build_excel_cli_eval, build_excel_mcp_eval):
                for supplied in (None, str(skill)):
                    with self.subTest(builder=builder.__name__, skill=supplied):
                        config = builder("offline", servers={}, skill_dir=supplied).build_session_config()
                        self.assertIs(config.get("enable_skills"), True)
                        self.assertFalse(config["enable_config_discovery"])
                        self.assertEqual(config.get("skill_directories", []), [supplied] if supplied else [])

    def test_cli_checks_reject_an_incomplete_final_execution(self):
        result = CopilotResult(turns=[Turn("assistant", "", [
            ToolCall("excel_execute", {"args": "--help"}),
        ])])
        with self.assertRaises(AssertionError):
            assert_cli_exit_codes(result)

    def test_cli_checks_reject_a_failed_batch_subcommand(self):
        result = CopilotResult(turns=[Turn("assistant", "", [
            ToolCall("excel_execute", {"args": "batch"}, result=(
                '{"exit_code":0,"stdout":"{\\"success\\":false,\\"errorMessage\\":\\"bad sheet\\"}\\n"}'
            ), completion_received=True, success=True),
        ])])
        with self.assertRaises(AssertionError):
            assert_cli_exit_codes(result)


if __name__ == "__main__":
    unittest.main()
