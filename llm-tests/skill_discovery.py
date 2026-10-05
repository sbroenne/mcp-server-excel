"""Verify real isolated SDK skill availability without sending model messages."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Any
from unittest.mock import patch

from copilot import CopilotClient, CopilotSession, PermissionHandler
from copilot.generated.rpc import ToolsListRequest
from pytest_skill_engineering import load_skill
from pytest_skill_engineering.copilot import CopilotEval


async def discover_skills(
    agent: CopilotEval, *, required_tools: tuple[str, ...] = (),
) -> list[dict[str, Any]]:
    config = agent.build_session_config()
    for directory in agent.skill_directories or []:
        load_skill(directory)
    with (
        patch.object(CopilotSession, "send", side_effect=AssertionError("Discovery must not send model requests")),
        patch.object(CopilotSession, "send_and_wait", side_effect=AssertionError("Discovery must not send model requests")),
        tempfile.TemporaryDirectory(prefix="skill-discovery-") as storage,
    ):
        client = CopilotClient(
            mode=agent.client_mode, base_directory=storage, working_directory=agent.working_directory,
            github_token=os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN"),
            log_level="warning",
        )
        try:
            await client.start()
            config["mcp_servers"] = {}
            config["on_permission_request"] = PermissionHandler.approve_all
            session = await client.create_session(**config)
            await session.rpc.skills.ensure_loaded()
            result = await session.rpc.skills.list()
            if required_tools:
                tools = (await client.rpc.tools.list(ToolsListRequest(model=agent.model))).tools
                names = {tool.name for tool in tools}
                assert set(required_tools) <= names, f"Missing native tools: {set(required_tools) - names}"
            return [{"name": skill.name, "enabled": skill.enabled, "path": str(skill.path)}
                    for skill in result.skills]
        finally:
            await client.stop()


def assert_skill_exposure(
    skills: list[dict[str, Any]], expected: str | None, directory: str | None = None,
) -> None:
    assert [skill["name"] for skill in skills] == ([expected] if expected else []), (
        f"Wrong skill availability: expected {expected or 'none'}, discovered {skills}"
    )
    assert all(skill["enabled"] is True for skill in skills), f"Requested skill is disabled: {skills}"
    if directory:
        assert len(skills) == 1 and Path(skills[0]["path"]).resolve() == (Path(directory) / "SKILL.md").resolve(), (
            f"Wrong skill source: expected {directory}, discovered {skills}"
        )
