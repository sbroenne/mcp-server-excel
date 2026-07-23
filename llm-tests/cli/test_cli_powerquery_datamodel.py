"""CLI Power Query and Data Model workflows."""

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
async def test_cli_star_schema_workflow(copilot_eval, excel_cli_servers, excel_cli_skill_dir, fixtures_dir):
    agent = build_excel_cli_eval(
        "cli-star-schema",
        servers=excel_cli_servers,
        skill_dir=excel_cli_skill_dir,
        max_turns=30,
    )

    products_json = (fixtures_dir / "products-dimension.json").as_posix()
    orders_json = (fixtures_dir / "orders-fact.json").as_posix()

    prompt = f"""
Using the Excel CLI tool, create a new workbook at {unique_path('star-schema-cli')}.

Build a complete star-schema workflow:

1. Create a Products sheet and write the JSON data from this file to A1:
{products_json}

2. Create an Orders sheet and write the JSON data from this file to A1:
{orders_json}

IMPORTANT:
- JSON arrays with commas break CLI argument parsing.
- You must use --values-file for both datasets instead of inline --values.

Then:
- convert the Products range into a table named Products and add it to the Data Model,
- convert the Orders range into a table named Orders and add it to the Data Model,
- create a relationship between Orders[ProductID] and Products[ProductID],
- create a PivotTable on a new sheet showing product category rows with sum of Quantity as values,
- add a pie chart showing quantity distribution by category,
- save the workbook.

Report:
- whether Electronics or Furniture has more total quantity,
- confirmation that both tables were added to the Data Model,
- confirmation that the relationship, PivotTable, and chart were created.
"""

    result = await copilot_eval(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
    assert_cli_args_contain(result, "--values-file")
    assert_regex(result.final_response, r"(?i)(electronics|furniture)")
    assert_regex(result.final_response, r"(?i)(relationship|pivot|chart|data model)")


@pytest.mark.asyncio
async def test_cli_powerquery_products_workflow(copilot_eval, excel_cli_servers, excel_cli_skill_dir, fixtures_dir):
    agent = build_excel_cli_eval(
        "cli-pq-products",
        servers=excel_cli_servers,
        skill_dir=excel_cli_skill_dir,
        max_turns=35,
    )

    mcode_file = (fixtures_dir / "products-powerquery.m").as_posix()

    prompt = f"""
Using the Excel CLI tool, create a new workbook at {unique_path('products-analysis-cli')}.

Use Power Query to create a query named Products from this M code file:
{mcode_file}

IMPORTANT:
- M code contains characters that break CLI inline parsing.
- You must use --mcode-file with the path above instead of inline --mcode.

Then:
- load the Products query to a worksheet as a table,
- add the loaded Products table to the Data Model,
- explain which columns are dimensions and which are facts/measures,
- create DAX measures for Average Rating, Total Products, Average Price, and Total Revenue,
- create a PivotTable on a new sheet with category rows and Total Products plus Average Rating as values,
- add a bar chart for that analysis,
- save the workbook.

Report:
- confirmation that the Power Query import succeeded,
- confirmation that the table is in the Data Model,
- which category has the most products,
- a short summary of the measures you created,
- confirmation that the chart was created and the file was saved.
"""

    result = await copilot_eval(agent, prompt)
    assert result.success
    assert_cli_exit_codes(result)
    assert_cli_args_contain(result, "--mcode-file")
    assert_regex(result.final_response, r"(?i)(dimension|fact|measure|data model)")
    assert_regex(result.final_response, r"(?i)(chart|saved)")
