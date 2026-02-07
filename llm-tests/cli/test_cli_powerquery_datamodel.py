"""CLI Power Query and Data Model workflows."""

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
async def test_cli_star_schema_workflow(aitest_run, excel_cli_server, excel_cli_skill, fixtures_dir):
    agent = Agent(
        name="cli-star-schema",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        max_turns=20,
    )

    products_json = (fixtures_dir / "products-dimension.json").as_posix()
    orders_json = (fixtures_dir / "orders-fact.json").as_posix()

    messages = None

    prompt = f"""
I need to set up a proper star schema for analysis.

Create a new Excel file at {unique_path('star-schema-cli')}

Create a sheet called "Products" and write the data from this JSON file to range A1:
{products_json}

IMPORTANT: JSON arrays with commas break CLI argument parsing.
You MUST use --values-file with the path above instead of --values with inline JSON.

Then make it a table called "Products" and add it to the Data Model.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_cli_args_contain(result, "--values-file")
    messages = result.messages

    prompt = f"""
Now let's add the transaction data.

Create a new sheet called "Orders" and write the data from this JSON file to range A1:
{orders_json}

IMPORTANT: JSON arrays with commas break CLI argument parsing.
You MUST use --values-file with the path above.

Then make it a table called "Orders" and add it to the Data Model.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_cli_args_contain(result, "--values-file")
    messages = result.messages

    prompt = """
Now link the tables together.

Create a relationship between the Orders and Products tables using ProductID.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    messages = result.messages

    prompt = """
Now for the analysis! Create a PivotTable on a new sheet that shows:
- Product Categories as rows (from the Products table)
- Sum of Quantity as the values (from the Orders table)
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(pivot|electronics|furniture|category)")
    messages = result.messages

    prompt = """
Add a pie chart showing the quantity distribution by category.

Save and close the file.

Which category had more orders - Electronics or Furniture?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(pie|chart|saved|closed|success)")


@pytest.mark.asyncio
async def test_cli_powerquery_products_workflow(
    aitest_run, excel_cli_server, excel_cli_skill, fixtures_dir
):
    agent = Agent(
        name="cli-pq-products",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        max_turns=20,
    )

    mcode_file = (fixtures_dir / "products-powerquery.m").as_posix()

    messages = None

    prompt = f"""
I want to analyze product sales data using Power Query.

Create a new Excel file at {unique_path('products-analysis-cli')}

Use Power Query to create a query named "Products" using the M code from this file:
{mcode_file}

IMPORTANT: M code contains special characters that break CLI parsing.
You MUST use --mcode-file with the path above. Do NOT try inline --mcode.

Load the query to a worksheet (the default behavior).
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_cli_args_contain(result, "--mcode-file")
    messages = result.messages

    prompt = """
The Products data from the Power Query is on a worksheet as an Excel Table, but I need it in Power Pivot (the Data Model) for DAX analysis.

Add the Products table to the Data Model so I can create DAX measures on it.

After adding to the Data Model, analyze the data structure:
- Which columns are dimensions (descriptive attributes for slicing/filtering)?
- Which columns are facts/measures (numeric values to aggregate)?

Confirm the table is now in the Data Model and ready for DAX measures.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(dimension|fact|data.?model|added)")
    messages = result.messages

    prompt = """
Now create some useful DAX measures on your Products table in the Data Model:

1. Average Rating: AVERAGE(Products[Rating])
2. Total Products: COUNTROWS(Products)
3. Average Price: AVERAGE(Products[Price])
4. Total Revenue: SUM(Products[Price])
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(measure|rating|price|revenue|created)")
    messages = result.messages

    prompt = """
Create a PivotTable on a new sheet using your star schema.

Show product categories as rows with Total Products and Average Rating as values.

Which category has the most products and what's their average rating?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(pivot|category|rating|products)")
    messages = result.messages

    prompt = """
Add a bar chart showing:
- Categories on the X-axis
- Total Products and Average Rating as values

Save and close the file.

Summarize the star schema you built: how many dimension tables, fact tables, relationships, and measures did you create?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(chart|star.?schema|dimension|fact|relationship|measure|saved|closed)")
