"""MCP Power Query and Data Model workflows."""

from __future__ import annotations

import pytest

from conftest import build_excel_mcp_eval, assert_regex, unique_results_path

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_star_schema_workflow(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    agent = build_excel_mcp_eval(
        "mcp-star-schema",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        allowed_tools=[
            "table",
            "datamodel",
            "datamodel_relationship",
            "pivottable",
            "chart",
            "chart_config",
            "range",
            "file",
            "worksheet",
        ],
        max_turns=30,
    )

    prompt = f"""
Create a new Excel workbook at {unique_results_path('star-schema-mcp')}.

Build a complete star-schema analysis workflow in one pass:

1. On a sheet named Products, enter:
   ProductID, ProductName, Category, UnitPrice
   P001, Laptop Pro, Electronics, 1200
   P002, Wireless Headphones, Electronics, 150
   P003, Standing Desk, Furniture, 400
   P004, Ergonomic Chair, Furniture, 250
   P005, USB Hub, Electronics, 35
   Convert the range to a table named Products and add it to the Data Model.

2. On a sheet named Orders, enter:
   OrderID, ProductID, Quantity, OrderDate
   1001, P001, 2, 2024-03-01
   1002, P002, 5, 2024-03-02
   1003, P003, 1, 2024-03-03
   1004, P001, 1, 2024-03-04
   1005, P004, 3, 2024-03-05
   1006, P002, 4, 2024-03-06
   1007, P005, 10, 2024-03-07
   1008, P001, 2, 2024-03-08
   Convert the range to a table named Orders and add it to the Data Model.

3. Create a relationship Orders[ProductID] -> Products[ProductID].

4. Create a PivotTable on a new sheet named Analysis that shows:
   - Category as rows,
   - sum of Quantity as values.

5. Add a pie chart showing quantity distribution by category.

Save the workbook and report:
- which category had more total quantity,
- how many tables were added to the Data Model,
- confirmation that the relationship, PivotTable, and chart were created.
"""

    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("table")
    assert result.tool_was_called("datamodel_relationship")
    assert result.tool_was_called("pivottable")
    assert result.tool_was_called("chart")
    assert_regex(result.final_response, r"(?i)(electronics|furniture)")
    assert_regex(result.final_response, r"(?i)(relationship|pivot|chart|data model)")


@pytest.mark.asyncio
async def test_mcp_powerquery_amazon_workflow(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir, fixtures_dir):
    agent = build_excel_mcp_eval(
        "mcp-amazon-pq",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        max_turns=35,
    )

    amazon_csv = (fixtures_dir / "amazon.csv").as_posix()

    prompt = f"""
Create a new Excel workbook at {unique_results_path('amazon-analysis-mcp')}.

Use Power Query to import the CSV file at:
{amazon_csv}

Name the query Products and load it to a worksheet as a table.

Then complete the analysis workflow:
1. Add the loaded Products table to the Data Model.
2. Explain which columns are dimensions and which are numeric facts/measures.
3. Create these DAX measures on the Products table:
   - Average Rating
   - Total Products
   - Average Discount Percentage
   - Total Potential Revenue (sum of original prices)
4. Create a PivotTable on a new sheet with Category as rows and Total Products plus Average Rating as values.
5. Add a bar chart based on that PivotTable.
6. Save the workbook.

Report:
- confirmation that the Power Query import succeeded,
- confirmation that the table is in the Data Model,
- which category has the most products,
- a short summary of the measures you created,
- confirmation that the chart was created and the file was saved.
"""

    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("powerquery")
    assert result.tool_was_called("datamodel")
    assert result.tool_was_called("pivottable")
    assert result.tool_was_called("chart")
    assert_regex(result.final_response, r"(?i)(dimension|fact|measure|data model)")
    assert_regex(result.final_response, r"(?i)(chart|saved)")
