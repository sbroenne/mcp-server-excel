"""CLI chart workflows."""

from __future__ import annotations

import pytest

from conftest import build_excel_cli_eval, assert_cli_exit_codes, assert_regex, unique_path

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_chart_workflows(copilot_eval, excel_cli_servers, excel_cli_skill_dir):
    agent = build_excel_cli_eval(
        "cli-chart-workflows",
        servers=excel_cli_servers,
        skill_dir=excel_cli_skill_dir,
        max_turns=30,
    )

    prompt = f"""
Using the Excel CLI tool, create a new workbook at {unique_path('chart-workflows-cli')}.

Build four chart scenarios in one workbook and save it at the end.

Scenario 1 - Table-backed column chart:
- On a sheet named SalesTable, enter:
  Product, Q1 Sales, Q2 Sales
  Laptop, 45000, 52000
  Phone, 38000, 41000
  Tablet, 22000, 28000
  Monitor, 15000, 18000
- Convert the range to a table named SalesData.
- Create a clustered column chart from SalesData.
- Position it below the data so it does not overlap rows 1-5.

Scenario 2 - Line chart below the source range:
- On a sheet named BudgetTrend, enter:
  Month, Revenue, Expenses
  January, 50000, 35000
  February, 55000, 38000
  March, 48000, 32000
  April, 62000, 41000
  May, 58000, 39000
- Create a line chart and position it below row 6.
- Read the chart details and confirm it starts at row 7 or later.

Scenario 3 - Two non-overlapping dashboard charts:
- On a sheet named MarketDashboard, enter:
  Company, Revenue, Market Share
  Alpha, 500000, 35
  Beta, 400000, 28
  Gamma, 300000, 22
  Delta, 200000, 15
- Convert the data to a table named MarketData.
- Create a pie chart for Market Share.
- Create a bar chart for Revenue.
- Confirm both charts exist and do not overlap.

Scenario 4 - Target-range positioning:
- On a sheet named QuarterlyPosition, enter:
  Region, Q1, Q2, Q3, Q4
  North, 1000, 1200, 1100, 1400
  South, 800, 900, 950, 1000
  East, 1500, 1600, 1450, 1700
  West, 600, 700, 650, 800
- Create a bar chart positioned near G2 using explicit chart positioning.

Save the workbook and report:
- that four charts were created,
- which chart types were used,
- confirmation that the dashboard charts do not overlap,
- confirmation that the BudgetTrend chart starts at row 7 or later,
- confirmation that the QuarterlyPosition chart is near G2.
"""

    result = await copilot_eval(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(4 charts|four charts)")
    assert_regex(result.final_response, r"(?i)(column|line|pie|bar)")
    assert_regex(result.final_response, r"(?i)(row 7|row seven|G2|near G2|no overlap|do not overlap)")
