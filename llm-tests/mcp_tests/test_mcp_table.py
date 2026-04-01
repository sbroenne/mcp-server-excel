"""MCP table workflows."""

from __future__ import annotations

import pytest

from conftest import (
    build_excel_mcp_eval,
    assert_regex,
    unique_path,
)

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_table_create_query(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    agent = build_excel_mcp_eval(
        "mcp-table-create",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        max_turns=20,
    )

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-table')}
2. On Sheet1, put these column headers in A1:D1: Product, Quantity, Price, Total
3. Add data in A2:D3:
   Row 2: Widget, 10, 5.99, 59.90
   Row 3: Gadget, 5, 12.99, 64.95
4. Create an Excel table from A1:D3 and name it "SalesData"
5. List all tables to confirm SalesData exists
6. Get the data from the SalesData table
7. Close the file without saving
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel-mcp-table")
    assert_regex(result.final_response, r"(?i)(SalesData)")


@pytest.mark.asyncio
async def test_mcp_table_lifecycle(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    agent = build_excel_mcp_eval(
        "mcp-table-lifecycle",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        max_turns=20,
    )

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-table-lifecycle')}
2. Put these column headers in A1:C1: ID, Name, Status
3. Add data in A2:C3:
   Row 2: 1, Task One, Active
   Row 3: 2, Task Two, Complete
4. Create a table from A1:C3 called "TaskList"
5. List all tables to verify TaskList was created
6. Delete the TaskList table
7. Close the file without saving
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel-mcp-table")
    assert_regex(result.final_response, r"(?i)(TaskList)")
