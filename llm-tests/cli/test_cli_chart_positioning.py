"""CLI chart positioning checked in the saved workbook."""

from __future__ import annotations

import pytest

from conftest import build_excel_cli_eval, assert_cli_exit_codes, unique_results_path
from workbook_assertions import assert_chart, read_saved_workbook
from workbook_scenarios import chart_prompt

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.cli]


@pytest.mark.asyncio
@pytest.mark.parametrize("below", [True, False], ids=["below-data", "right-of-table"])
async def test_cli_chart_position(copilot_eval, excel_cli_servers, excel_cli_skill_dir, below):
    path = unique_results_path("chart-cli")
    agent = build_excel_cli_eval(
        f"cli-chart-{'below' if below else 'right'}",
        servers=excel_cli_servers, skill_dir=excel_cli_skill_dir,
    )
    result = await copilot_eval(agent, chart_prompt(path, below=below))
    assert result.success
    assert_cli_exit_codes(result)
    assert_chart(read_saved_workbook(path, "A1:C6" if below else "A1:D5"), below=below)
