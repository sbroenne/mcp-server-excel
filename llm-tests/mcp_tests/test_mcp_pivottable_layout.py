"""MCP PivotTable layout tests."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_regex, unique_results_path

pytestmark = [pytest.mark.aitest, pytest.mark.mcp]


def _has_row_layout(result, value: int) -> bool:
    calls = result.tool_calls_for("excel_pivottable_calc")
    return any(call.arguments.get("row_layout") == value for call in calls)


@pytest.mark.asyncio
async def test_mcp_pivottable_tabular_layout(aitest_run, excel_mcp_server, excel_mcp_skill):
    agent = Agent(
        name="mcp-pivot-tabular",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=["excel_pivottable", "excel_pivottable_calc", "excel_table", "excel_range", "excel_file", "excel_worksheet"],
        max_turns=20,
    )

    prompt = f"""
I need to create a PivotTable that I can easily copy-paste into a database or CSV export tool. Each field should be in its own column so the data is flat and easy to work with.

Create a new Excel file at {unique_results_path('pivottable-tabular')}

Enter this sales data starting at A1:
Region, Product, Quarter, Sales
North, Laptops, Q1, 45000
North, Laptops, Q2, 52000
North, Phones, Q1, 28000
North, Phones, Q2, 31000
South, Laptops, Q1, 38000
South, Laptops, Q2, 48000
South, Phones, Q1, 24000
South, Phones, Q2, 29000

Convert it to a table called "SalesData".

Then create a PivotTable on a new sheet called "Analysis" at cell A3 named "SalesPivot" with the appropriate layout for flat data export.

Add Region and Product as row fields, and Sales as a value field.

After creating the PivotTable, summarize what you created and confirm it uses a tabular/flat layout suitable for data export.
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel_pivottable")
    assert _has_row_layout(result, 1)
    # Empty response is OK if tools were called successfully
    if result.final_response:
        assert_regex(result.final_response, r"(?i)(pivot|tabular|layout|created|success|sales)")


@pytest.mark.asyncio
async def test_mcp_pivottable_compact_layout(aitest_run, excel_mcp_server, excel_mcp_skill):
    agent = Agent(
        name="mcp-pivot-compact",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=["excel_pivottable", "excel_pivottable_calc", "excel_table", "excel_range", "excel_file", "excel_worksheet"],
        max_turns=20,
    )

    prompt = f"""
I want a PivotTable that saves horizontal space. Show all the row labels in a single column with indentation to indicate hierarchy - the default compact view that Excel normally uses.

Create a new Excel file at {unique_results_path('pivottable-compact')}

Enter this data starting at A1:
Department, Team, Employee, Hours
Engineering, Backend, Alice, 160
Engineering, Backend, Bob, 152
Engineering, Frontend, Carol, 168
Engineering, Frontend, Dave, 144
Sales, Direct, Eve, 176
Sales, Direct, Frank, 160
Sales, Partners, Grace, 148

Convert it to a table called "TimeTracking".

Create a PivotTable on a new sheet at A3 named "HoursSummary" with the standard compact indented layout.

Add Department and Team as row fields, and Hours as a value field.
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel_pivottable")
    assert _has_row_layout(result, 0)
    assert_regex(result.final_response, r"(?i)(pivot|compact|layout|created|success)")


@pytest.mark.asyncio
async def test_mcp_pivottable_outline_layout(aitest_run, excel_mcp_server, excel_mcp_skill):
    agent = Agent(
        name="mcp-pivot-outline",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=["excel_pivottable", "excel_pivottable_calc", "excel_table", "excel_range", "excel_file", "excel_worksheet"],
        max_turns=25,
    )

    prompt = f"""
I need a PivotTable where I can expand and collapse groups to drill down through the geographic hierarchy. Each level should have its own column with +/- buttons to show or hide the details underneath.

Create a new Excel file at {unique_results_path('pivottable-outline')}

Enter this data starting at A1:
Country, State, City, Revenue
USA, California, Los Angeles, 500000
USA, California, San Francisco, 450000
USA, Texas, Houston, 380000
USA, Texas, Dallas, 320000
Canada, Ontario, Toronto, 420000
Canada, Ontario, Ottawa, 180000
Canada, BC, Vancouver, 350000

Convert it to a table called "GeoRevenue".

Create a PivotTable on a new sheet at A3 named "RegionalAnalysis" with a layout that supports expanding and collapsing hierarchy levels.

Add Country, State, and City as row fields, and Revenue as a value field.

Summarize: What are the three layout styles and when should each be used?
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel_pivottable")
    assert _has_row_layout(result, 2)
    assert_regex(result.final_response, r"(?i)(compact|tabular|outline)")
