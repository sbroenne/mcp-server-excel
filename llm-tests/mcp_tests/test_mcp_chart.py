"""MCP chart workflows."""

from __future__ import annotations

import pytest

from conftest import build_excel_mcp_eval, assert_regex, unique_path

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_chart_workflows(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    agent = build_excel_mcp_eval(
        "mcp-chart-workflows",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        allowed_tools=["chart", "chart_config", "table", "file", "range", "worksheet"],
        max_turns=30,
    )

    prompt = f"""
Create a new Excel workbook at {unique_path('chart-workflows-mcp')}.

Build four independent chart scenarios in the same workbook and save it at the end.

Scenario 1 - Table-backed column chart:
- On a sheet named "SalesTable", enter:
  Product, Q1 Sales, Q2 Sales
  Laptop, 45000, 52000
  Phone, 38000, 41000
  Tablet, 22000, 28000
  Monitor, 15000, 18000
- Convert the data to a table named SalesData.
- Create a clustered column chart from SalesData.
- Position it below the data so the chart does not overlap rows 1-5.

Scenario 2 - Line chart below the source range:
- On a sheet named "BudgetTrend", enter:
  Month, Revenue, Expenses
  January, 50000, 35000
  February, 55000, 38000
  March, 48000, 32000
  April, 62000, 41000
  May, 58000, 39000
- Create a line chart from the data.
- Position it below row 6 and confirm the chart starts at row 7 or later.

Scenario 3 - Two non-overlapping dashboard charts:
- On a sheet named "MarketDashboard", enter:
  Company, Revenue, Market Share
  Alpha, 500000, 35
  Beta, 400000, 28
  Gamma, 300000, 22
  Delta, 200000, 15
- Convert the data to a table named MarketData.
- Create a pie chart for Market Share.
- Create a bar chart for Revenue.
- Place the two charts so they do not overlap each other or the data table.

Scenario 4 - Target-range positioning:
- On a sheet named "QuarterlyPosition", enter:
  Region, Q1, Q2, Q3, Q4
  North, 1000, 1200, 1100, 1400
  South, 800, 900, 950, 1000
  East, 1500, 1600, 1450, 1700
  West, 600, 700, 650, 800
- Create a bar chart and place it near G2 using an explicit target range.

Save the workbook and report:
- how many charts were created,
- which chart types were used,
- confirmation that the dashboard charts do not overlap,
- confirmation that the BudgetTrend chart starts at row 7 or later,
- confirmation that the QuarterlyPosition chart is near G2.
"""

    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("chart")
    assert result.tool_was_called("table")
    assert_regex(result.final_response, r"(?i)(4 charts|four charts)")
    assert_regex(result.final_response, r"(?i)(column|line|pie|bar)")
    assert_regex(result.final_response, r"(?i)(row 7|row seven|G2|near G2|no overlap|do not overlap)")
