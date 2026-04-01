"""MCP styling workflows — validates correct style system usage per object type."""

from __future__ import annotations

import pytest

from conftest import (
    build_excel_mcp_eval,
    assert_regex,
    unique_path,
)

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_styling_table_style(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    """LLM should use table(set-style) for table visual styling, not range_format on header."""
    agent = build_excel_mcp_eval(
        "mcp-styling-table",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        max_turns=20,
    )

    prompt = f"""
Create a new Excel file at {unique_path('llm-test-styling-table')}

Enter this quarterly sales data on Sheet1:
Region, Q1, Q2, Q3, Q4
North, 120000, 135000, 118000, 142000
South, 98000, 102000, 115000, 128000
West, 85000, 91000, 99000, 108000

Format the data as a professional Excel Table named "QuarterlySales"
with a visually appealing style.

Close the file without saving.
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel-mcp-table")
    assert_regex(result.final_response, r"(?i)(QuarterlySales|table|style)")


@pytest.mark.asyncio
async def test_mcp_styling_semantic_status(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    """LLM should use range_format(set-style) with Good/Bad/Neutral for status cells."""
    agent = build_excel_mcp_eval(
        "mcp-styling-status",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        max_turns=20,
    )

    prompt = f"""
Create a new Excel file at {unique_path('llm-test-styling-status')}

Enter this project status data on Sheet1:
Task, Owner, Status
Design, Alice, Complete
Development, Bob, In Progress
Testing, Carol, Overdue
Deployment, Dave, Complete

Format the Status column cells with distinct colours to make the status
visually clear at a glance — green for Complete, red for Overdue,
yellow or neutral for In Progress.

Close the file without saving.
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert_regex(result.final_response, r"(?i)(format|style|colour|color|green|red|conditional)")


@pytest.mark.asyncio
async def test_mcp_styling_header_fill(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    """LLM should use format-range (not set-style) for a header row with a fill colour."""
    agent = build_excel_mcp_eval(
        "mcp-styling-header",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        max_turns=20,
    )

    prompt = f"""
Create a new Excel file at {unique_path('llm-test-styling-header')}

Enter this data on Sheet1:
Product, Units, Revenue
Widget, 450, 13500
Gadget, 280, 19600
Doohickey, 175, 8750

Give the header row (row 1) a dark blue background with white bold text,
centred horizontally.

Close the file without saving.
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("excel-mcp-range_format")
    assert_regex(result.final_response, r"(?i)(header|format|blue|white|bold)")
