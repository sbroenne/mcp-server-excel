"""Real SDK discovery checks; no Excel or paid model messages."""

from pathlib import Path

import pytest

from conftest import build_excel_cli_eval, build_excel_mcp_eval
from skill_discovery import assert_skill_exposure, discover_skills


@pytest.mark.parametrize("transport", ("mcp", "cli"))
@pytest.mark.parametrize("with_skill", (False, True))
async def test_skill_availability_without_model_messages(
    github_auth, tmp_path: Path, excel_mcp_skill_dir, excel_cli_skill_dir, transport, with_skill,
):
    builder = build_excel_mcp_eval if transport == "mcp" else build_excel_cli_eval
    supplied = excel_mcp_skill_dir if transport == "mcp" else excel_cli_skill_dir
    agent = builder("discovery-only", servers={}, working_directory=str(tmp_path),
                    skill_dir=supplied if with_skill else None)
    discovered = await discover_skills(agent)
    assert_skill_exposure(discovered, f"excel-{transport}-report-formatting" if with_skill else None,
                          supplied if with_skill else None)
