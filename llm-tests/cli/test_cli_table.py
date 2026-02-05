"""CLI table workflows."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_cli_exit_codes, assert_regex, unique_path

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_table_create_query(aitest_run, excel_cli_server, excel_cli_skill):
    agent = Agent(
        name="cli-table-create",
        provider=Provider(model="azure/gpt-5-mini", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        max_turns=20,
    )

    prompt = f"""
Using the Excel CLI tool:
1. Create a new empty Excel file at {unique_path('llm-test-table-cli')}
2. On Sheet1, put these column headers in A1:D1: Product, Quantity, Price, Total
3. Add data in A2:D3:
   Row 2: Widget, 10, 5.99, 59.90
   Row 3: Gadget, 5, 12.99, 64.95
4. Create an Excel table from A1:D3 and name it "SalesData"
5. List all tables to confirm SalesData exists
6. Get the data from the SalesData table
7. Close the file without saving
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(SalesData)")


@pytest.mark.asyncio
async def test_cli_table_lifecycle(aitest_run, excel_cli_server, excel_cli_skill):
    agent = create_cli_agent(excel_cli_server, excel_cli_skill, name="cli-table-lifecycle")

    prompt = f"""
Using the Excel CLI tool:
1. Create a new empty Excel file at {unique_path('llm-test-table-lifecycle-cli')}
2. Put these column headers in A1:C1: ID, Name, Status
3. Add data in A2:C3:
   Row 2: 1, Task One, Active
   Row 3: 2, Task Two, Complete
4. Create a table from A1:C3 called "TaskList"
5. List all tables to verify TaskList was created
6. Delete the TaskList table
7. Close the file without saving
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(TaskList)")
