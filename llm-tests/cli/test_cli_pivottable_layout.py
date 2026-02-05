"""CLI PivotTable layout tests."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_cli_exit_codes, assert_regex, unique_path

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_pivottable_tabular_layout(aitest_run, excel_cli_server, excel_cli_skill):
    agent = Agent(
        name="cli-pivot-tabular",
        provider=Provider(model="azure/gpt-5-mini", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        max_turns=20,
    )

    prompt = f"""
I need to create a PivotTable that I can easily copy-paste into a database or CSV export tool. Each field should be in its own column so the data is flat.

Create a new Excel file at {unique_path('pivottable-tabular-cli')}

Enter this sales data starting at A1:
Region, Product, Quarter, Sales
North, Laptops, Q1, 45000
North, Laptops, Q2, 52000
North, Phones, Q1, 28000
North, Phones, Q2, 31000
South, Laptops, Q1, 38000
South, Laptops, Q2, 48000
South, Phones, Q1, 24000
South, Phones, Q2, 29000

Convert it to a table called "SalesData".

Then create a PivotTable on a new sheet called "Analysis" with Tabular layout.
Add Region and Product as row fields, and Sales as a value field.

Save and close the file.
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(pivot|tabular|created|success)")


@pytest.mark.asyncio
async def test_cli_pivottable_compact_layout(aitest_run, excel_cli_server, excel_cli_skill):
    agent = create_cli_agent(excel_cli_server, excel_cli_skill, name="cli-pivot-compact")

    prompt = f"""
I want a PivotTable with the default compact view that Excel normally uses - row labels in a single column.

Create a new Excel file at {unique_path('pivottable-compact-cli')}

Enter this data starting at A1:
Department, Team, Employee, Hours
Engineering, Backend, Alice, 160
Engineering, Backend, Bob, 152
Engineering, Frontend, Carol, 168
Sales, Direct, Eve, 176
Sales, Partners, Grace, 148

Convert it to a table called "TimeTracking".

Create a PivotTable on a new sheet with Compact layout.
Add Department and Team as row fields, and Hours as a value field.

Save and close the file.
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(pivot|compact|created|success)")


@pytest.mark.asyncio
async def test_cli_pivottable_outline_layout(aitest_run, excel_cli_server, excel_cli_skill):
    agent = create_cli_agent(excel_cli_server, excel_cli_skill, name="cli-pivot-outline")

    prompt = f"""
I need a PivotTable with Outline layout for expanding and collapsing groups.

Create a new Excel file at {unique_path('pivottable-outline-cli')}

Enter this data starting at A1:
Country, State, City, Revenue
USA, California, Los Angeles, 500000
USA, California, San Francisco, 450000
USA, Texas, Houston, 380000
Canada, Ontario, Toronto, 420000
Canada, BC, Vancouver, 350000

Convert it to a table called "GeoRevenue".

Create a PivotTable on a new sheet with Outline layout.
Add Country, State, and City as row fields, and Revenue as a value field.

Save and close the file.

Summarize: What are the three PivotTable layout styles?
"""
    result = await aitest_run(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(compact|tabular|outline)")
