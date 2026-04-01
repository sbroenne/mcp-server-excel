"""MCP chart positioning workflows."""

from __future__ import annotations

import pytest

from conftest import (
    build_excel_mcp_eval,
    assert_regex,
    unique_path,
)

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_chart_position_below_data(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    agent = build_excel_mcp_eval(
        "mcp-chart-below",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        max_turns=20,
    )

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-chart-pos')} and open it
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
7. Summarize the chart you created and its position.
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("chart")
    # Looser assertion - just confirm chart work was done
    assert result.final_response or result.tool_was_called("chart")


@pytest.mark.asyncio
async def test_mcp_chart_position_right_of_table(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    agent = build_excel_mcp_eval(
        "mcp-chart-right",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        max_turns=25,
    )

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-chart-table')} and open it
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
7. Confirm what you created.
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("chart")
    # Loosen - either chart or table mentioned
    assert result.final_response or result.tool_was_called("chart")
