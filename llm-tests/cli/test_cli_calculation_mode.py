"""CLI calculation mode workflow."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_cli_exit_codes, assert_regex, unique_path, DEFAULT_RETRIES, DEFAULT_TIMEOUT_MS

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.xfail(reason="LLM may not autonomously use calculation_mode for small batches", strict=False)
@pytest.mark.asyncio
async def test_cli_calculation_mode_batch_flow(aitest_run, excel_cli_server, excel_cli_skill):
    agent = Agent(
        name="cli-calc-mode",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
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
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(manual|automatic|calculation)")
