"""CLI calculation mode workflow."""

from __future__ import annotations

import pytest

from conftest import (
    build_excel_cli_eval,
    assert_cli_exit_codes,
    assert_regex,
    unique_path,
)

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_calculation_mode_batch_flow(copilot_eval, excel_cli_servers, excel_cli_skill_dir):
    agent = build_excel_cli_eval(
        "cli-calc-mode",
        servers=excel_cli_servers,
        skill_dir=excel_cli_skill_dir,
        max_turns=20,
    )

    prompt = f"""
Create a new Excel file at {unique_path('calc-mode-cli')}

Set calculation mode to manual.

On Sheet1, write this data in A1:C4:
Category, Budget, Actual
Rent, 1000, 1000
Food, 500, 450
Transport, 200, 180

Add a formula in D2:D4 for Variance = C2-B2, C3-B3, C4-B4.

After all writes, explicitly recalculate the workbook.

Switch calculation mode back to automatic.

Report the current calculation mode and the variance values.
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(manual|automatic|calculation)")
