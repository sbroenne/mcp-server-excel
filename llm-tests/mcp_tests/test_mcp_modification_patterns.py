"""MCP modification pattern tests (targeted updates)."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_regex, unique_path

pytestmark = [pytest.mark.aitest, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_range_updates(aitest_run, excel_mcp_server, excel_mcp_skill):
    agent = Agent(
        name="mcp-range-updates",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=["excel_range", "excel_file", "excel_worksheet"],
        max_turns=20,
    )

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-range')} and open it
2. Set up a budget in A1:C4 on Sheet1:
   Row 1: Category, Budget, Actual
   Row 2: Rent, 1000, 1000
   Row 3: Food, 500, 450
   Row 4: Transport, 200, 180
3. Put a canary formula =ROW()*1000+COLUMN() in cell D1 (will show 1004)
4. Update Food actual (C3) to 480
5. Update Transport actual (C4) to 195
6. Add a new row 5: Utilities, 150, 145
7. Read D1 to verify the canary formula still shows 1004
8. Read all data from A1:C5 to verify the updates
9. Close the file without saving
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel_range")
    assert_regex(result.final_response, r"(?i)(1004)")
    assert_regex(result.final_response, r"(?i)(480)")
    assert_regex(result.final_response, r"(?i)(195)")


@pytest.mark.asyncio
async def test_mcp_table_updates(aitest_run, excel_mcp_server, excel_mcp_skill):
    agent = Agent(
        name="mcp-table-updates",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=["excel_range", "excel_table", "excel_file", "excel_worksheet"],
        max_turns=20,
    )

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-table')} and open it
2. Put sales data in A1:C4 on Sheet1:
   Row 1: Product, Price, Quantity
   Row 2: Widget, 25, 3
   Row 3: Gadget, 50, 2
   Row 4: Device, 75, 1
3. Convert range A1:C4 into an Excel Table named "SalesTable"
4. Add column header "Total" in D1, then formulas =B2*C2 in D2, =B3*C3 in D3, =B4*C4 in D4
5. Update Widget quantity (C2) from 3 to 5 - this should NOT delete the table
6. List tables to confirm exactly 1 table named "SalesTable" exists
7. Read D2 to verify formula recalculated (should be 125)
8. Close the file without saving
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel_table")
    assert_regex(result.final_response, r"(?i)(salestable)")
    assert_regex(result.final_response, r"(?i)(125)")


@pytest.mark.asyncio
async def test_mcp_chart_updates(aitest_run, excel_mcp_server, excel_mcp_skill):
    agent = Agent(
        name="mcp-chart-updates",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=["excel_chart", "excel_file", "excel_worksheet", "excel_range"],
        max_turns=20,
    )

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-chart')} and open it
2. Put chart data in A1:B4 on Sheet1:
   Row 1: Month, Sales
   Row 2: Jan, 100
   Row 3: Feb, 150
   Row 4: Mar, 200
3. Create a column chart from A1:B4 with title "Monthly Sales"
4. Change ONLY the chart title to "Q1 Sales Report" - do NOT delete and recreate the chart
5. List charts to confirm exactly 1 chart exists
6. Read chart info to verify the title is now "Q1 Sales Report"
7. Close the file without saving
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel_chart")
    assert_regex(result.final_response, r"(?i)(q1 sales report)")


@pytest.mark.asyncio
async def test_mcp_sheet_structural_changes(aitest_run, excel_mcp_server, excel_mcp_skill):
    agent = Agent(
        name="mcp-sheet-struct",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=["excel_range", "excel_file", "excel_worksheet"],
        max_turns=20,
    )

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-struct')} and open it
2. Put employee data in A1:C5 on Sheet1:
   Row 1: Name, Department, ID
   Row 2: Alice, Engineering, =ROW()*100 (formula shows 200)
   Row 3: Bob, Marketing, =ROW()*100 (formula shows 300)
   Row 4: Carol, Sales, =ROW()*100 (formula shows 400)
   Row 5: Dave, Support, =ROW()*100 (formula shows 500)
3. Delete the row with "Bob" (row 3), shifting remaining rows up
4. Rename the "Department" header (B1) to "Team"
5. Read all data including column C to verify:
   - Alice's ID formula now shows 200
   - Carol's ID formula now shows 300
   - Dave's ID formula now shows 400
   - Bob is gone
   - Header B1 says "Team"
6. Close the file without saving
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel_range")
    assert_regex(result.final_response, r"(?i)(200)")
    assert_regex(result.final_response, r"(?i)(300)")
    assert_regex(result.final_response, r"(?i)(400)")
    assert_regex(result.final_response, r"(?i)(team)")
