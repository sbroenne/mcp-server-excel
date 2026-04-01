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
        allowed_tools=["file", "range", "chart", "screenshot"],
        max_turns=20,
    )


@pytest.mark.asyncio
async def test_mcp_targetrange_no_skill(copilot_eval, excel_mcp_servers):
    """targetRange parameter should work without skill guidance."""
    agent = build_excel_mcp_eval(
        "targetrange-no-skill",
        servers=excel_mcp_servers,
        allowed_tools=["file", "range", "chart", "screenshot"],
        max_turns=20,
    )
    assert_regex(result.final_response, r"(?i)(chart|created|F2|position)")


@pytest.mark.asyncio
async def test_mcp_multi_chart_collision_no_skill(
    copilot_eval, excel_mcp_servers,
):
    """Multi-chart dashboard should avoid overlaps using built-in collision detection, no skill."""
    agent = build_excel_mcp_eval(
        "multi-chart-collision-no-skill",
        servers=excel_mcp_servers,
        allowed_tools=["file", "range", "chart", "screenshot"],
        max_turns=25,
    )


@pytest.mark.asyncio
async def test_mcp_collision_warning_reaction_no_skill(copilot_eval, excel_mcp_servers):
    """LLM should react to OVERLAP WARNING by repositioning, without skill guidance."""
    agent = build_excel_mcp_eval(
        "collision-reaction-no-skill",
        servers=excel_mcp_servers,
        allowed_tools=["file", "range", "chart", "screenshot"],
        max_turns=25,
    )
    # LLM should mention overlap/warning/reposition in its summary
    assert_regex(result.final_response, r"(?i)(overlap|warning|reposition|move|fix)")
