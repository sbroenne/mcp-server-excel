"""CLI file and worksheet workflows."""

from __future__ import annotations

import pytest

from conftest import (
    build_excel_cli_eval,
    assert_cli_exit_codes,
    unique_path,
)

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_file_and_worksheet_workflow(copilot_eval, excel_cli_servers, excel_cli_skill_dir):
    agent = build_excel_cli_eval(
        "cli-file-worksheet",
        servers=excel_cli_servers,
        skill_dir=excel_cli_skill_dir,
        max_turns=20,
    )

    prompt = f"""
Create a new Excel file at {unique_path('budget')}

Set it up with two sheets: Income and Expenses.

On the Income sheet, add this data starting at A1:
- Headers: Source, Amount
- Salary: 5000
- Freelance: 1200

On the Expenses sheet, add:
- Headers: Category, Amount
- Rent: 1500
- Utilities: 200
- Food: 600

Save the file when done.
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
