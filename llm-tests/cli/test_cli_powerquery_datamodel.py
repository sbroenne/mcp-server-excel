"""CLI Power Query and Data Model workflows."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import (
    assert_cli_args_contain,
    assert_cli_exit_codes,
    assert_regex,
    unique_path,
    DEFAULT_RETRIES,
    DEFAULT_TIMEOUT_MS,
)

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_star_schema_workflow(aitest_run, excel_cli_server, excel_cli_skill, fixtures_dir):
    agent = Agent(
        name="cli-star-schema",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        max_turns=30,
        retries=DEFAULT_RETRIES,
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
    result = await aitest_run(agent, prompt, messages=messages, timeout_ms=DEFAULT_TIMEOUT_MS)
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
    result = await aitest_run(agent, prompt, messages=messages, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_cli_args_contain(result, "--values-file")
    messages = result.messages

    prompt = """
Now link the tables together.

Create a relationship between the Orders and Products tables using ProductID.
"""
    result = await aitest_run(agent, prompt, messages=messages, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    messages = result.messages

    prompt = """
Now for the analysis! Create a PivotTable on a new sheet that shows:
- Product Categories as rows (from the Products table)
- Sum of Quantity as the values (from the Orders table)
"""
    result = await aitest_run(agent, prompt, messages=messages, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(pivot|electronics|furniture|category)")
    messages = result.messages

    prompt = """
Add a pie chart showing the quantity distribution by category.

Save and close the file.

Which category had more orders - Electronics or Furniture?
"""
    result = await aitest_run(agent, prompt, messages=messages, timeout_ms=DEFAULT_TIMEOUT_MS)
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
        retries=DEFAULT_RETRIES,
    )

    mcode_file = (fixtures_dir / "products-powerquery.m").as_posix()

    messages = None

    prompt = f"""
I want to analyze product sales data using Power Query.

Create a new Excel file at {unique_path('products-analysis-cli')}

Use Power Query to create a query named "Products" using the M code from this file:
{mcode_file}

Load the query to a worksheet.

Confirm the query was created and how many rows of data were loaded.
"""
    result = await aitest_run(agent, prompt, messages=messages, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_cli_args_contain(result, "--mcode-file")
    messages = result.messages

    prompt = """
Add the Products table to the Data Model.

Then create a PivotTable on a new sheet showing product categories as rows
with a count of products as the value.

Save and close the file.

Summarize which category has the most products.
"""
    result = await aitest_run(agent, prompt, messages=messages, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(pivot|category|products|data.?model)")
