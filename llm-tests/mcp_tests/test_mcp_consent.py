"""MCP permission guidance checked through calls and workbook outcomes."""

from __future__ import annotations

import pytest

from conftest import build_excel_mcp_eval
from consent_scenarios import (
    SCENARIOS, assert_consent_outcome, consent_prompt, consent_workbook, record_questions,
)

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]


@pytest.mark.asyncio
@pytest.mark.parametrize("scenario", SCENARIOS)
async def test_mcp_consent(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir, consent_workbook, scenario):
    questions = []
    agent = record_questions(build_excel_mcp_eval(
        f"mcp-consent-{scenario}", servers=excel_mcp_servers, skill_dir=excel_mcp_skill_dir,
    ), questions)
    result = await copilot_eval(agent, consent_prompt(consent_workbook.path, scenario))
    assert_consent_outcome(result, consent_workbook, scenario, questions)
