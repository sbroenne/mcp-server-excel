"""CLI range workflows."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import (
    assert_cli_args_contain,
    assert_cli_exit_codes,
    assert_regex,
    unique_path,
)

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_range_set_get(aitest_run, excel_cli_server, excel_cli_skill, fixtures_dir):
    agent = Agent(
        name="cli-range",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
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
    result = await aitest_run(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
    assert_cli_args_contain(result, "--values-file")
    assert_regex(result.final_response, r"(?i)(Product)")


@pytest.mark.asyncio
async def test_cli_range_error_handling(aitest_run, excel_cli_server, excel_cli_skill):
    agent = create_cli_agent(excel_cli_server, excel_cli_skill, name="cli-range-error")

    prompt = f"""
1. Create a new empty Excel file at {unique_path('llm-test-range-error-cli')}
2. Try to get values from a large range like A1:Z1000 to see what happens
3. Then close the file without saving
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
