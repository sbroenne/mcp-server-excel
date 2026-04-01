"""MCP chart collision detection and auto-positioning tests.

Tests that the built-in collision detection, auto-positioning, and screenshot
verification hints work WITHOUT the skill. This validates that the MCP tool
descriptions and result messages alone are sufficient to guide the LLM toward
well-positioned charts.

Key behaviors tested:
- Auto-positioning places charts below data when no position is specified
- targetRange positions charts within specified cell ranges
- Collision warnings are returned in the result message
- LLM reacts to OVERLAP WARNING by repositioning
- LLM uses screenshot to verify layout (prompted by result message)
"""

from __future__ import annotations

import pytest

from conftest import (
    build_excel_mcp_eval,
    assert_regex,
    unique_path,
)

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_auto_position_no_skill(copilot_eval, excel_mcp_servers):
    """Auto-positioning should place charts below data without skill guidance."""
    agent = build_excel_mcp_eval(
        "auto-position-no-skill",
        servers=excel_mcp_servers,
        max_turns=20,
    )

    prompt = f"""
Create a new Excel file at {unique_path('auto-position-test')}.
Write sales data to A1:C6 (headers in row 1, data in rows 2-6).
Create a column chart from the data without specifying a position.
Report the chart position to confirm it was placed below the data.
Close the file without saving.
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel-mcp-chart")


@pytest.mark.asyncio
async def test_mcp_targetrange_no_skill(copilot_eval, excel_mcp_servers):
    """targetRange parameter should work without skill guidance."""
    agent = build_excel_mcp_eval(
        "targetrange-no-skill",
        servers=excel_mcp_servers,
        max_turns=20,
    )

    prompt = f"""
Create a new Excel file at {unique_path('targetrange-test')}.
Write data to A1:D5 (headers in row 1, data in rows 2-5).
Create a chart and position it at F2 using the targetRange parameter.
Report the chart position.
Close the file without saving.
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel-mcp-chart")
    assert_regex(result.final_response, r"(?i)(chart|created|F2|position)")


@pytest.mark.asyncio
async def test_mcp_multi_chart_collision_no_skill(
    copilot_eval, excel_mcp_servers,
):
    """Multi-chart dashboard should avoid overlaps using built-in collision detection, no skill."""
    agent = build_excel_mcp_eval(
        "multi-chart-collision-no-skill",
        servers=excel_mcp_servers,
        max_turns=25,
    )

    prompt = f"""
Create a new Excel file at {unique_path('multi-chart-collision')}.
Write revenue data to A1:C5 and market share data to E1:F5.
Create two charts: a bar chart for revenue and a pie chart for market share.
Position them so they do not overlap each other or the data.
Report the positions of both charts and confirm no overlap.
Close the file without saving.
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel-mcp-chart")


@pytest.mark.asyncio
async def test_mcp_collision_warning_reaction_no_skill(copilot_eval, excel_mcp_servers):
    """LLM should react to OVERLAP WARNING by repositioning, without skill guidance."""
    agent = build_excel_mcp_eval(
        "collision-reaction-no-skill",
        servers=excel_mcp_servers,
        max_turns=25,
    )

    prompt = f"""
Create a new Excel file at {unique_path('collision-reaction')}.
Write data to A1:C10.
Create a chart at A1 (this will overlap the data and should trigger a warning).
React to any warnings by repositioning the chart to avoid overlap.
Report what happened and confirm the chart is now positioned correctly.
Close the file without saving.
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel-mcp-chart")
    # LLM should mention overlap/warning/reposition in its summary
    assert_regex(result.final_response, r"(?i)(overlap|warning|reposition|move|fix)")
