"""MCP chart positioning checked in the saved workbook."""

from __future__ import annotations

import pytest

from conftest import build_excel_mcp_eval, unique_results_path
from workbook_assertions import assert_chart, read_saved_workbook
from workbook_scenarios import chart_prompt

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]


@pytest.mark.asyncio
@pytest.mark.parametrize("below", [True, False], ids=["below-data", "right-of-table"])
async def test_mcp_chart_position(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir, below):
    path = unique_results_path("chart-mcp")
    agent = build_excel_mcp_eval(
        f"mcp-chart-{'below' if below else 'right'}",
        servers=excel_mcp_servers, skill_dir=excel_mcp_skill_dir,
    )
    result = await copilot_eval(agent, chart_prompt(path, below=below))
    assert result.success
    assert result.tool_was_called("excel-mcp-chart")
    assert_chart(read_saved_workbook(path, "A1:C6" if below else "A1:D5"), below=below)
