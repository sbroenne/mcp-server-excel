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

from conftest import (
    build_excel_mcp_eval,
    assert_regex,
    unique_results_path,
)

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]


# =============================================================================
# Tests WITH Skill - LLM guided by skill documentation
# =============================================================================


@pytest.mark.asyncio
async def test_mcp_calculation_mode_batch_with_skill(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    """Test that LLM uses manual calculation mode for batch writes (with skill).

    The skill provides guidance on when to use calculation mode.
    """
    agent = build_excel_mcp_eval(
        "mcp-calc-batch-skill",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        allowed_tools=[
            "calculation_mode",
            "file",
            "range",
            "worksheet",
        ],
        max_turns=25,
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
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("calculation_mode"), \
        "LLM with skill should use calculation_mode for batch writes"
    assert result.tool_was_called("range")
    assert_regex(result.final_response, r"(?i)(total|grand|sum|\d{4,})")


# =============================================================================
# Tests WITHOUT Skill - LLM relies purely on tool descriptions
# =============================================================================


@pytest.mark.asyncio
async def test_mcp_calculation_mode_batch_no_skill(copilot_eval, excel_mcp_servers):
    """Test that LLM uses manual calculation mode for batch writes (no skill).

    Without the skill, the LLM must discover the calculation mode tool
    purely from its description. This tests tool discoverability.
    """
    agent = build_excel_mcp_eval(
        "mcp-calc-batch-noskill",
        servers=excel_mcp_servers,
        allowed_tools=[
            "calculation_mode",
            "file",
            "range",
            "worksheet",
        ],
        max_turns=25,
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
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("calculation_mode"), \
        "LLM without skill should discover and use calculation_mode for batch writes"
    assert result.tool_was_called("range")
    assert_regex(result.final_response, r"(?i)(total|grand|sum|\d{4,})")

