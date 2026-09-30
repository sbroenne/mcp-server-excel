"""MCP slicer workflows checked after reopening the saved workbook."""

from __future__ import annotations

import pytest

from conftest import build_excel_mcp_eval, unique_results_path
from workbook_assertions import (
    assert_combined_slicers, assert_pivot_slicer, assert_table_slicers, read_saved_workbook,
)
from workbook_scenarios import combined_slicer_prompt, pivot_slicer_prompt, table_slicer_prompt

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]


@pytest.mark.asyncio
@pytest.mark.parametrize("prompt,check", [
    (pivot_slicer_prompt, assert_pivot_slicer),
    (table_slicer_prompt, assert_table_slicers),
    (combined_slicer_prompt, assert_combined_slicers),
], ids=["pivot", "table", "combined"])
async def test_mcp_slicer_workflow(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir, prompt, check):
    path = unique_results_path("slicer-mcp")
    agent = build_excel_mcp_eval(
        f"mcp-{prompt.__name__}", servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir, max_turns=30,
    )
    result = await copilot_eval(agent, prompt(path))
    assert result.success
    assert result.tool_was_called("excel-mcp-slicer")
    check(read_saved_workbook(path))
