"""CLI chart positioning workflows."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_cli_exit_codes, assert_regex, unique_path

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_chart_position_below_data(aitest_run, excel_cli_server, excel_cli_skill):
    agent = Agent(
        name="cli-chart-below",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        max_turns=20,
    )

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-chart-pos-cli')} and open it
2. Put sales data in A1:C6 on Sheet1:
   Row 1: Month, Revenue, Expenses
   Row 2: January, 50000, 35000
   Row 3: February, 55000, 38000
   Row 4: March, 48000, 32000
   Row 5: April, 62000, 41000
   Row 6: May, 58000, 39000
3. Create a column chart from B1:C6 with Month labels from A1:A6
4. Position the chart so it does NOT overlap with the data - it should be placed BELOW row 6
5. List the charts and report the exact chart position
6. Save and close the file
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(chart|created)")


@pytest.mark.asyncio
async def test_cli_chart_position_right_of_table(aitest_run, excel_cli_server, excel_cli_skill):
    agent = create_cli_agent(excel_cli_server, excel_cli_skill, name="cli-chart-right")

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-chart-table-cli')} and open it
2. Put product data in A1:D5 on Sheet1:
   Row 1: Product, Q1, Q2, Q3
   Row 2: Widget, 100, 150, 120
   Row 3: Gadget, 80, 90, 110
   Row 4: Device, 200, 180, 220
   Row 5: Tool, 50, 60, 75
3. Convert A1:D5 into an Excel Table named "ProductSales"
4. Create a line chart from the table's numeric data (columns B:D)
5. Position the chart to the RIGHT of the table so it doesn't overlap
6. Save and close the file
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(chart|created|productsales)")
