"""CLI complete sales report workflow."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_cli_exit_codes, assert_regex, unique_path, DEFAULT_RETRIES, DEFAULT_TIMEOUT_MS

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_sales_report_workflow(aitest_run, excel_cli_server, excel_cli_skill):
    agent = Agent(
        name="cli-sales-report",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    messages = None

    prompt = f"""
Create a new Excel file at {unique_path('sales-report-cli')}

Enter this sales data starting at A1:
Region, Product, Salesperson, Revenue
North, Laptop, Alice, 12000
South, Monitor, Bob, 8500
East, Keyboard, Carol, 3600
West, Mouse, Dave, 2500
North, Monitor, Alice, 9000
South, Laptop, Bob, 15000

Convert the data (A1:D7) into an Excel Table called "SalesTransactions".

Confirm the table name and how many data rows it has.
"""
    result = await aitest_run(agent, prompt, messages=messages, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(salestransactions)")
    assert_regex(result.final_response, r"(?i)(6|six)")
    messages = result.messages

    prompt = """
Add 2 more rows to the SalesTransactions table:
East, Laptop, Carol, 14400
West, Monitor, Dave, 7200

Confirm the table now has 8 data rows and save the file.
"""
    result = await aitest_run(agent, prompt, messages=messages, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(8|eight)")
