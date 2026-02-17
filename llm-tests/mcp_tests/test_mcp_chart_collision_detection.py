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

from pytest_aitest import Agent, Provider

from conftest import assert_regex, unique_path, DEFAULT_RETRIES, DEFAULT_TIMEOUT_MS

pytestmark = [pytest.mark.aitest, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_auto_position_no_skill(aitest_run, excel_mcp_server):
    """Auto-positioning should place charts below data without skill guidance."""
    agent = Agent(
        name="auto-position-no-skill",
        provider=Provider(model="azure/gpt-4.1"),
        mcp_servers=[excel_mcp_server],
        # NO skill — only tool descriptions guide the LLM
        allowed_tools=["file", "range", "chart", "screenshot"],
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
1. Create a new Excel file at {unique_path('auto-pos-no-skill')} and open it
2. Put this data in Sheet1 starting at A1:
   Month, Sales, Cost
   Jan, 50000, 35000
   Feb, 55000, 38000
   Mar, 48000, 32000
   Apr, 62000, 41000
3. Create a column chart from the data. Do NOT specify left, top, or targetRange - let the server auto-position it.
4. Save and close the file.
5. Report the chart position and whether there were any overlap warnings.
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert result.tool_was_called("chart")
    # The LLM should follow the result message hint and verify with screenshot
    assert result.tool_was_called("screenshot"), (
        "Expected LLM to use screenshot for verification (prompted by result message)"
    )


@pytest.mark.asyncio
async def test_mcp_targetrange_no_skill(aitest_run, excel_mcp_server):
    """targetRange parameter should work without skill guidance."""
    agent = Agent(
        name="targetrange-no-skill",
        provider=Provider(model="azure/gpt-4.1"),
        mcp_servers=[excel_mcp_server],
        # NO skill
        allowed_tools=["file", "range", "chart", "screenshot"],
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
1. Create a new Excel file at {unique_path('targetrange-no-skill')} and open it
2. Put quarterly revenue data in Sheet1 A1:D4:
   Region, Q1, Q2, Q3
   North, 1000, 1200, 1100
   South, 800, 900, 950
   East, 1500, 1600, 1450
3. Create a bar chart from the data. Position it within cell range F2:L15 using the targetRange parameter.
4. Save and close.
5. Report the chart name and position.
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert result.tool_was_called("chart")
    assert result.tool_was_called("screenshot"), (
        "Expected LLM to use screenshot for verification (prompted by result message)"
    )
    assert_regex(result.final_response, r"(?i)(chart|created|F2|position)")


@pytest.mark.asyncio
async def test_mcp_multi_chart_collision_no_skill(
    aitest_run, excel_mcp_server,
):
    """Multi-chart dashboard should avoid overlaps using built-in collision detection, no skill."""
    agent = Agent(
        name="multi-chart-collision-no-skill",
        provider=Provider(model="azure/gpt-4.1"),
        mcp_servers=[excel_mcp_server],
        # NO skill — relies on tool descriptions + collision warnings
        # Minimal tool set to reduce schema size and token usage
        allowed_tools=["file", "range", "chart", "screenshot"],
        max_turns=25,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
1. Create a new Excel file at {unique_path('multi-collision-no-skill')} and open it
2. Enter sales data in Sheet1 A1:C5:
   Product, Revenue, Units
   Laptop, 450000, 1500
   Phone, 380000, 8000
   Tablet, 220000, 3000
   Monitor, 150000, 2000
3. Create a clustered column chart from A1:B5 (Revenue by Product). Let the server auto-position it.
4. Create a pie chart from A1:A5,C1:C5 (Units by Product). Use targetRange to place it to the right of the first chart.
   If you get any overlap warnings, fix the positions.
5. Save and close the file.
6. Summarize the dashboard layout including both chart positions.
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS * 2)
    assert result.success
    assert result.tool_was_called("chart")
    assert result.tool_was_called("screenshot"), (
        "LLM must take screenshot to verify multi-chart layout has no overlaps"
    )


@pytest.mark.asyncio
async def test_mcp_collision_warning_reaction_no_skill(aitest_run, excel_mcp_server):
    """LLM should react to OVERLAP WARNING by repositioning, without skill guidance."""
    agent = Agent(
        name="collision-reaction-no-skill",
        provider=Provider(model="azure/gpt-4.1"),
        mcp_servers=[excel_mcp_server],
        # NO skill
        allowed_tools=["file", "range", "chart", "screenshot"],
        max_turns=25,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
1. Create a new Excel file at {unique_path('collision-react-no-skill')} and open it
2. Put data in Sheet1 A1:B5:
   Category, Value
   Alpha, 100
   Beta, 200
   Gamma, 150
   Delta, 175
3. Create a column chart from A1:B5, but deliberately place it at left=0, top=0 (which should overlap the data).
4. If there's an OVERLAP WARNING in the result, fix the chart position so it no longer overlaps.
5. Save and close.
6. Tell me whether there was an overlap warning and how you fixed it.
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert result.tool_was_called("chart")
    assert result.tool_was_called("screenshot"), (
        "Expected LLM to use screenshot for verification (prompted by result message)"
    )
    # LLM should mention overlap/warning/reposition in its summary
    assert_regex(result.final_response, r"(?i)(overlap|warning|reposition|move|fix)")
