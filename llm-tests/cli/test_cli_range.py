"""CLI range workflows."""

from __future__ import annotations

import pytest

from conftest import (
    build_excel_cli_eval,
    assert_cli_args_contain,
    assert_cli_exit_codes,
    assert_regex,
    unique_path,
)

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_range_set_get(copilot_eval, excel_cli_servers, excel_cli_skill_dir, fixtures_dir):
    agent = build_excel_cli_eval(
        "cli-range",
        servers=excel_cli_servers,
        skill_dir=excel_cli_skill_dir,
        max_turns=20,
    )
    values_file = (fixtures_dir / "range-test-data.json").as_posix()

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-range-cli')}
2. Write data to Sheet1 range A1:C2 using the values from this JSON file:
   {values_file}

   IMPORTANT: JSON arrays with commas break CLI argument parsing.
   You MUST use --values-file with the path above instead of --values with inline JSON.
3. Read back the data from A1:C2 to verify it was written correctly
4. Close the file without saving
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
    assert_cli_args_contain(result, "--values-file")
    assert_regex(result.final_response, r"(?i)(Product)")


@pytest.mark.asyncio
async def test_cli_range_error_handling(copilot_eval, excel_cli_servers, excel_cli_skill_dir):
    agent = build_excel_cli_eval(
        "cli-range-error",
        servers=excel_cli_servers,
        skill_dir=excel_cli_skill_dir,
        max_turns=20,
    )

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-range-error-cli')}
2. Try to get values from a large range like A1:Z1000 to see what happens
3. Then close the file without saving
"""
    result = await copilot_eval(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
