"""CLI slicer workflows."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_cli_exit_codes, assert_regex, unique_path

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_pivottable_slicer_workflow(aitest_run, excel_cli_server, excel_cli_skill):
    agent = Agent(
        name="cli-pivot-slicer",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        max_turns=20,
    )

    messages = None

    prompt = f"""
I want to test PivotTable slicers.

Create a new Excel file at {unique_path('pivottable-slicer-cli')}

On Sheet1, enter this sales data starting at A1:

Region, Product, Quarter, Sales
North, Laptop, Q1, 15000
North, Phone, Q1, 8000
North, Laptop, Q2, 18000
North, Phone, Q2, 9500
South, Laptop, Q1, 12000
South, Phone, Q1, 7500
South, Laptop, Q2, 14000
South, Phone, Q2, 8200

Convert this to a table called "SalesData".

Then create a PivotTable on a new sheet called "Analysis" that shows:
- Region as rows
- Sum of Sales as values
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(pivot|region|sales|created)")
    messages = result.messages

    prompt = """
Create a slicer for the Region field on the PivotTable.

After creating, list all slicers to confirm it was created.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(slicer|region|created|success)")
    messages = result.messages

    prompt = """
Use the Region slicer to show only "North" region data.

After applying the filter, what does the PivotTable show for total North sales?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(north|filter|slicer|50500|sales)")
    messages = result.messages

    prompt = """
Delete the slicer we created.

Save and close the file.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(delete|removed|closed|saved|success)")


@pytest.mark.asyncio
async def test_cli_table_slicer_workflow(aitest_run, excel_cli_server, excel_cli_skill):
    agent = Agent(
        name="cli-table-slicer",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        max_turns=20,
    )

    messages = None

    prompt = f"""
I want to test Table slicers (different from PivotTable slicers).

Create a new Excel file at {unique_path('table-slicer-cli')}

On Sheet1, enter this employee data starting at A1:

Department, Employee, Status, Salary
Engineering, Alice, Active, 85000
Engineering, Bob, Active, 92000
Marketing, Carol, Active, 78000
Marketing, Dave, Inactive, 70000
Sales, Eve, Active, 65000
Sales, Frank, Inactive, 62000

Convert this to an Excel table called "Employees".
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(employees|table|created|success)")
    messages = result.messages

    prompt = """
Create a Table slicer for the Department column.

List all Table slicers to confirm it was created.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(slicer|department|table|created|success)")
    messages = result.messages

    prompt = """
Use the Department slicer to filter the table to show only Engineering employees.

How many Engineering employees are there?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(engineering|2|two|filter|alice|bob)")
    messages = result.messages

    prompt = """
Delete the Table slicer.

Save and close the file.

What's the difference between Table slicers and PivotTable slicers?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(delete|table|pivot|different|closed|saved)")
