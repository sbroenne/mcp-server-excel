"""MCP calculation mode workflows.

Tests that the LLM autonomously uses calculation_mode for batch performance
optimization - should recognize bulk writes benefit from manual mode.

Tests both WITH skill (guided by skill documentation) and WITHOUT skill
(relying purely on tool descriptions) to ensure discoverability.

Note: Formula verification tests were removed because manual mode is NOT required
to read formula text - get-formulas works in any calculation mode.
"""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_regex, unique_results_path, DEFAULT_RETRIES, DEFAULT_TIMEOUT_MS

pytestmark = [pytest.mark.aitest, pytest.mark.mcp]


# =============================================================================
# Tests WITH Skill - LLM guided by skill documentation
# =============================================================================


@pytest.mark.asyncio
@pytest.mark.xfail(reason="LLM may not autonomously use calculation_mode for small batches", strict=False)
async def test_mcp_calculation_mode_batch_with_skill(aitest_run, excel_mcp_server, excel_mcp_skill):
    """Test that LLM uses manual calculation mode for batch writes (with skill).

    The skill provides guidance on when to use calculation mode.
    """
    agent = Agent(
        name="mcp-calc-batch-skill",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=[
            "calculation_mode",
            "file",
            "range",
            "worksheet",
        ],
        max_turns=25,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
Build a sales summary worksheet with the following data.

Create a new Excel file at {unique_results_path('calc-batch-skill')}

On the first sheet, enter:
- A1: "Product", B1: "Price", C1: "Qty", D1: "Total"
- A2: "Laptop", B2: 1200, C2: 5
- A3: "Monitor", B3: 450, C3: 8
- A4: "Keyboard", B4: 120, C4: 20
- A5: "Mouse", B5: 25, C5: 50

Add formulas in column D to calculate totals (Price * Qty) for rows 2-5.
Add a grand total formula in D6 that sums D2:D5.

Report the calculated grand total in D6.
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert result.tool_was_called("calculation_mode"), \
        "LLM with skill should use calculation_mode for batch writes"
    assert result.tool_was_called("range")
    assert_regex(result.final_response, r"(?i)(total|grand|sum|\d{4,})")


# =============================================================================
# Tests WITHOUT Skill - LLM relies purely on tool descriptions
# =============================================================================


@pytest.mark.asyncio
@pytest.mark.xfail(reason="LLM may not autonomously use calculation_mode for small batches", strict=False)
async def test_mcp_calculation_mode_batch_no_skill(aitest_run, excel_mcp_server):
    """Test that LLM uses manual calculation mode for batch writes (no skill).

    Without the skill, the LLM must discover the calculation mode tool
    purely from its description. This tests tool discoverability.
    """
    agent = Agent(
        name="mcp-calc-batch-noskill",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        mcp_servers=[excel_mcp_server],
        # No skill - relying on tool descriptions only
        allowed_tools=[
            "calculation_mode",
            "file",
            "range",
            "worksheet",
        ],
        max_turns=25,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
Build a sales summary worksheet with the following data.

Create a new Excel file at {unique_results_path('calc-batch-noskill')}

On the first sheet, enter:
- A1: "Product", B1: "Price", C1: "Qty", D1: "Total"
- A2: "Laptop", B2: 1200, C2: 5
- A3: "Monitor", B3: 450, C3: 8
- A4: "Keyboard", B4: 120, C4: 20
- A5: "Mouse", B5: 25, C5: 50

Add formulas in column D to calculate totals (Price * Qty) for rows 2-5.
Add a grand total formula in D6 that sums D2:D5.

Report the calculated grand total in D6.
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert result.tool_was_called("calculation_mode"), \
        "LLM without skill should discover and use calculation_mode for batch writes"
    assert result.tool_was_called("range")
    assert_regex(result.final_response, r"(?i)(total|grand|sum|\d{4,})")

