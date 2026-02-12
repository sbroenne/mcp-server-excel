"""CLI file and worksheet workflows."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_cli_exit_codes, unique_path

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_file_and_worksheet_workflow(aitest_run, excel_cli_server, excel_cli_skill):
    agent = Agent(
        name="cli-file-worksheet",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
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
    result = await aitest_run(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
