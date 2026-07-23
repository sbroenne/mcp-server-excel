"""MCP range workflows."""

from __future__ import annotations

import pytest

from conftest import (
    build_excel_mcp_eval,
    assert_regex,
    unique_path,
)

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_range_set_get(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir, fixtures_dir):
    agent = build_excel_mcp_eval(
        "mcp-range",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        max_turns=20,
    )
    values_file = (fixtures_dir / "range-test-data.json").as_posix()

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-range')}
2. Write data to Sheet1 range A1:C2 using these values:
   Row 1: Product, Quantity, Price
   Row 2: Widget, 10, 5.99
3. Read back the data from A1:C2 to verify it was written correctly
4. Close the file without saving
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel-mcp-range")
    assert_regex(result.final_response, r"(?i)(Product)")


@pytest.mark.asyncio
async def test_mcp_range_error_handling(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    agent = build_excel_mcp_eval(
        "mcp-range-error",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        max_turns=20,
    )

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-range-error')}
2. Try to get values from a large range like A1:Z1000 to see what happens
3. Then close the file without saving
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
