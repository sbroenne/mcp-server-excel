"""CLI permission guidance checked through calls and workbook outcomes."""

from __future__ import annotations

import hashlib

import pytest

from conftest import build_excel_cli_eval
from consent_scenarios import (
    SCENARIOS, assert_consent_outcome, assert_read_only, consent_prompt, consent_workbook,
    isolated_cli_servers, record_questions,
)

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.cli]


@pytest.mark.asyncio
@pytest.mark.parametrize("scenario", SCENARIOS)
async def test_cli_consent(copilot_eval, excel_cli_servers, excel_cli_skill_dir, consent_workbook, scenario):
    questions = []
    agent = record_questions(build_excel_cli_eval(
        f"cli-consent-{scenario}",
        servers=isolated_cli_servers(excel_cli_servers, consent_workbook.pipe),
        skill_dir=excel_cli_skill_dir,
    ), questions)
    result = await copilot_eval(agent, consent_prompt(consent_workbook.path, scenario))
    assert_consent_outcome(result, consent_workbook, scenario, questions)


@pytest.mark.asyncio
async def test_cli_audit_preserves_existing_unsaved_session(
    copilot_eval, excel_cli_servers, excel_cli_skill_dir, consent_workbook,
):
    """CLI-only: its private daemon can have a user session before the agent starts."""
    workbook = consent_workbook
    session = workbook.cli("session", "open", str(workbook.path))["sessionId"]
    workbook.cli("calculationmode", "set-settings", "--session", session, "--mode", "manual")
    workbook.cli("range", "set-values", "--session", session, "--sheet", "Sheet1",
                 "--range", "B7", "--values", '[["Unsaved user note"]]')
    questions = []
    agent = record_questions(build_excel_cli_eval(
        "cli-audit-existing-session",
        servers=isolated_cli_servers(excel_cli_servers, workbook.pipe),
        skill_dir=excel_cli_skill_dir,
    ), questions)
    result = await copilot_eval(agent, consent_prompt(workbook.path, "audit"))
    assert result.success, result.error
    assert_read_only(result)
    assert not questions
    sessions = workbook.cli("session", "list")["sessions"]
    assert len(sessions) == 1 and sessions[0]["sessionId"] == session, sessions
    assert workbook.cli("range", "get-values", "--session", session, "--sheet", "Sheet1", "--range", "B7")["values"] == [["Unsaved user note"]]
    assert workbook.cli("range", "get-formulas", "--session", session, "--sheet", "Sheet1", "--range", "C5")["formulas"] == [["=SUM(C2:C3)"]]
    assert workbook.cli("calculationmode", "get-settings", "--session", session)["mode"] == "manual"
    assert workbook.cli("window", "get-info", "--session", session)["isVisible"] is False
    assert workbook.cli("workbook", "get-info", "--session", session)["saved"] is False
    assert hashlib.sha256(workbook.path.read_bytes()).hexdigest() == workbook.original_hash
