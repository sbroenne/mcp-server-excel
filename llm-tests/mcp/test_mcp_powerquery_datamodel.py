"""MCP Power Query and Data Model workflows."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_regex, unique_results_path

pytestmark = [pytest.mark.aitest, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_star_schema_workflow(aitest_run, excel_mcp_server, excel_mcp_skill):
    agent = Agent(
        name="mcp-star-schema",
        provider=Provider(model="azure/gpt-5-mini", rpm=10, tpm=10000),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=[
            "excel_table",
            "excel_datamodel",
            "excel_datamodel_rel",
            "excel_pivottable",
            "excel_chart",
            "excel_range",
            "excel_file",
            "excel_worksheet",
        ],
        max_turns=20,
    )

    messages = None

    prompt = f"""
I need to set up a proper star schema for analysis.

Create a new Excel file at {unique_results_path('star-schema')}

On a sheet called "Products", enter this product catalog starting at A1:

ProductID, ProductName, Category, UnitPrice
P001, Laptop Pro, Electronics, 1200
P002, Wireless Headphones, Electronics, 150
P003, Standing Desk, Furniture, 400
P004, Ergonomic Chair, Furniture, 250
P005, USB Hub, Electronics, 35

Make it a table called "Products" and add it to the Data Model.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_table")
    messages = result.messages

    prompt = """
Now let's add the transaction data.

Create a new sheet called "Orders" with this data starting at A1:

OrderID, ProductID, Quantity, OrderDate
1001, P001, 2, 2024-03-01
1002, P002, 5, 2024-03-02
1003, P003, 1, 2024-03-03
1004, P001, 1, 2024-03-04
1005, P004, 3, 2024-03-05
1006, P002, 4, 2024-03-06
1007, P005, 10, 2024-03-07
1008, P001, 2, 2024-03-08

Make it a table called "Orders" and add it to the Data Model.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    messages = result.messages

    prompt = """
Now link the tables together.

Create a relationship between the Orders and Products tables using ProductID.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_datamodel_rel")
    messages = result.messages

    prompt = """
Now for the analysis! Create a PivotTable on a new sheet that shows:
- Product Categories as rows (from the Products table)
- Sum of Quantity as the values (from the Orders table)
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_pivottable")
    messages = result.messages

    prompt = """
Add a pie chart showing the quantity distribution by category.

Save and close the file.

Which category had more orders - Electronics or Furniture?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_chart")
    assert_regex(result.final_response, r"(?i)(pie|chart|saved|closed|success)")


@pytest.mark.asyncio
async def test_mcp_powerquery_amazon_workflow(aitest_run, excel_mcp_server, excel_mcp_skill, fixtures_dir):
    agent = create_mcp_agent(excel_mcp_server, excel_mcp_skill, name="mcp-amazon-pq")

    messages = None
    amazon_csv = (fixtures_dir / "amazon.csv").as_posix()

    prompt = f"""
I want to analyze Amazon product sales data.

Create a new Excel file at {unique_results_path('amazon-analysis')}

Use Power Query to import this CSV file:
{amazon_csv}

Name the query "Products".
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_powerquery")
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
    assert result.tool_was_called("excel_table")
    assert_regex(result.final_response, r"(?i)(dimension|fact|data.?model|added)")
    messages = result.messages

    prompt = """
Now create some useful DAX measures on your fact table:

1. Average Rating
2. Total Products
3. Average Discount Percentage
4. Total Potential Revenue (sum of original prices)
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_datamodel")
    assert_regex(result.final_response, r"(?i)(measure|rating|discount|revenue|created)")
    messages = result.messages

    prompt = """
Create a PivotTable on a new sheet using your star schema.

Show product categories as rows with Total Products and Average Rating as values.

Which category has the most products and what's their average rating?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_pivottable")
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
    assert result.tool_was_called("excel_chart")
    assert_regex(result.final_response, r"(?i)(chart|star.?schema|dimension|fact|relationship|measure|saved|closed)")
