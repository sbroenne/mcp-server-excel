"""CLI styling workflows — validates correct style system usage per object type."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_cli_exit_codes, assert_regex, unique_path, DEFAULT_RETRIES, DEFAULT_TIMEOUT_MS

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_styling_table_style(aitest_run, excel_cli_server, excel_cli_skill):
    """LLM should use table(set-style) for table visual styling, not range_format on header."""
    agent = Agent(
        name="cli-styling-table",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
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
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(QuarterlySales|table|style)")


@pytest.mark.asyncio
async def test_cli_styling_semantic_status(aitest_run, excel_cli_server, excel_cli_skill):
    """LLM should use range_format(set-style) with Good/Bad/Neutral for status cells."""
    agent = Agent(
        name="cli-styling-status",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
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
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(format|style|colour|color|green|red)")


@pytest.mark.asyncio
async def test_cli_styling_header_fill(aitest_run, excel_cli_server, excel_cli_skill):
    """LLM should use format-range (not set-style) for a header row with a fill colour."""
    agent = Agent(
        name="cli-styling-header",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
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
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(header|format|blue|white|bold)")
