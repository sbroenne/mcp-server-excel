"""MCP chart workflows."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_regex, unique_path

pytestmark = [pytest.mark.aitest, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_chart_workflows(aitest_run, excel_mcp_server, excel_mcp_skill):
    agent = Agent(
        name="mcp-chart-workflows",
        provider=Provider(model="azure/gpt-5-mini", rpm=10, tpm=10000),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=["excel_chart", "excel_table", "excel_file", "excel_range", "excel_worksheet"],
        max_turns=20,
    )

    messages = None

    prompt = f"""
Create a sales analysis chart:

1. Create a new Excel file at {unique_path('chart-from-table')}
2. Enter this sales data in A1:C5:
   Product, Q1 Sales, Q2 Sales
   Laptop, 45000, 52000
   Phone, 38000, 41000
   Tablet, 22000, 28000
   Monitor, 15000, 18000
3. Convert the data to a table called "SalesData"
4. Create a column chart from the table using create-from-table
5. Position it below the data (targetRange='A8:F20')
6. Save and close
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_chart")
    assert result.tool_was_called("excel_table")
    assert_regex(result.final_response, r"(?i)(salesdata|chart|created)")
    messages = result.messages

    prompt = f"""
I need a chart that doesn't overlap my data:

1. Create a new Excel file at {unique_path('chart-position')}
2. Put this budget data in A1:C6:
   Month, Revenue, Expenses
   January, 50000, 35000
   February, 55000, 38000
   March, 48000, 32000
   April, 62000, 41000
   May, 58000, 39000
3. Create a line chart from the data
4. Make sure the chart is positioned BELOW row 6 so it doesn't cover the data
5. Read chart info and confirm the topLeftCell is row 7 or later
6. Save and close
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_chart")
    assert_regex(result.final_response, r"(?i)(chart|row [7-9]|\$[A-Z]+\$[7-9])")
    messages = result.messages

    prompt = f"""
Create a dashboard with multiple chart types:

1. Create a new Excel file at {unique_path('multi-chart')}
2. Enter market data in A1:C5:
   Company, Revenue, Market Share
   Alpha, 500000, 35
   Beta, 400000, 28
   Gamma, 300000, 22
   Delta, 200000, 15
3. Convert to a table called "MarketData"
4. Create a PIE chart showing Market Share, position at F2:K12
5. Create a BAR chart showing Revenue, position at F14:K24
6. List charts and confirm both exist without overlapping
7. Save and close
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_chart")
    assert_regex(result.final_response, r"(?i)(pie|bar|2 chart|two chart)")
    messages = result.messages

    prompt = f"""
Create a chart with precise cell-based positioning:

1. Create a new Excel file at {unique_path('chart-targetrange')}
2. Enter quarterly data in A1:E5:
   Region, Q1, Q2, Q3, Q4
   North, 1000, 1200, 1100, 1400
   South, 800, 900, 950, 1000
   East, 1500, 1600, 1450, 1700
   West, 600, 700, 650, 800
3. Create a bar chart using targetRange='G2:L15' for positioning
4. Verify the chart's topLeftCell is at or near G2
5. Save and close
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_chart")
    assert_regex(result.final_response, r"(?i)(\$G\$2|G2|chart|created)")
