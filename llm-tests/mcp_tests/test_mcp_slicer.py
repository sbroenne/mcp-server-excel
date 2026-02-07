"""MCP slicer workflows for PivotTables and Tables."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_regex, unique_results_path

pytestmark = [pytest.mark.aitest, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_pivottable_slicer_workflow(aitest_run, excel_mcp_server, excel_mcp_skill):
    agent = Agent(
        name="mcp-pivot-slicer",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=[
            "excel_pivottable",
            "excel_slicer",
            "excel_table",
            "excel_range",
            "excel_file",
            "excel_worksheet",
        ],
        max_turns=20,
    )

    messages = None

    prompt = f"""
I want to test PivotTable slicers. Let's set up the data first.

Create a new Excel file at {unique_results_path('pivottable-slicer')}

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
    assert result.tool_was_called("excel_pivottable")
    messages = result.messages

    prompt = """
Now I want to filter this PivotTable interactively.

Create a slicer for the Region field on the PivotTable we just made.
Position it at cell E2.

After creating, list all slicers to confirm it was created.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_slicer")
    assert_regex(result.final_response, r"(?i)(slicer|region|created|success)")
    messages = result.messages

    prompt = """
Great! Now use the Region slicer to show only "North" region data.

After applying the filter, what does the PivotTable show for total North sales?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_slicer")
    assert_regex(result.final_response, r"(?i)(north|filter|slicer|50500|sales)")
    messages = result.messages

    prompt = """
Now create a second slicer for the Product field, positioned at cell G2.

Then clear the Region filter so all regions show again.

How many slicers do we have now?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_slicer")
    assert_regex(result.final_response, r"(?i)(slicer|product|2|two|created)")
    messages = result.messages

    prompt = """
Delete both slicers we created.

Save and close the file.

Confirm both slicers were removed.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_file")
    assert_regex(result.final_response, r"(?i)(delete|removed|closed|saved|success)")


@pytest.mark.asyncio
async def test_mcp_table_slicer_workflow(aitest_run, excel_mcp_server, excel_mcp_skill):
    agent = Agent(
        name="mcp-table-slicer",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=[
            "excel_slicer",
            "excel_table",
            "excel_range",
            "excel_file",
            "excel_worksheet",
        ],
        max_turns=20,
    )

    messages = None

    prompt = f"""
I want to test Table slicers (different from PivotTable slicers).

Create a new Excel file at {unique_results_path('table-slicer')}

On Sheet1, enter this employee data starting at A1:

Department, Employee, Status, Salary
Engineering, Alice, Active, 85000
Engineering, Bob, Active, 92000
Marketing, Carol, Active, 78000
Marketing, Dave, Inactive, 70000
Sales, Eve, Active, 65000
Sales, Frank, Inactive, 62000
Engineering, Grace, Active, 88000
Sales, Henry, Active, 71000

Convert this to an Excel table called "Employees".
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_table")
    messages = result.messages

    prompt = """
Now create a Table slicer for the Department column.
Position it at cell F2.

List all Table slicers to confirm it was created.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_slicer")
    assert_regex(result.final_response, r"(?i)(slicer|department|table|created|success)")
    messages = result.messages

    prompt = """
Use the Department slicer to filter the table to show only Engineering employees.

How many Engineering employees are there?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_slicer")
    assert_regex(result.final_response, r"(?i)(engineering|3|three|filter|alice|bob|grace)")
    messages = result.messages

    prompt = """
Add another Table slicer for the Status column at cell H2.

Then filter to show only "Active" employees across all departments.

How many active employees are there total?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_slicer")
    assert_regex(result.final_response, r"(?i)(status|active|6|six|slicer)")
    messages = result.messages

    prompt = """
Delete all the Table slicers we created.

Save and close the file.

Summarize: what's the difference between Table slicers and PivotTable slicers?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_file")
    assert_regex(result.final_response, r"(?i)(delete|table|pivot|different|closed|saved)")


@pytest.mark.asyncio
async def test_mcp_combined_slicer_workflow(aitest_run, excel_mcp_server, excel_mcp_skill):
    agent = create_mcp_agent(excel_mcp_server, excel_mcp_skill, name="mcp-combined-slicer")

    messages = None

    prompt = f"""
Create a new Excel file at {unique_results_path('combined-slicer')}

On Sheet1, enter this inventory data starting at A1:

Category, Product, Warehouse, Stock, Price
Electronics, Laptop, West, 50, 999
Electronics, Phone, West, 120, 599
Electronics, Laptop, East, 35, 999
Electronics, Phone, East, 80, 599
Furniture, Desk, West, 25, 350
Furniture, Chair, West, 40, 175
Furniture, Desk, East, 30, 350
Furniture, Chair, East, 55, 175

1. Convert this to a table called "Inventory"
2. Create a PivotTable on a new sheet called "Summary" that shows:
   - Category as rows
   - Sum of Stock as values
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_table")
    assert result.tool_was_called("excel_pivottable")
    messages = result.messages

    prompt = """
I want to create slicers for both the Table and the PivotTable.

1. On Sheet1, create a TABLE slicer for the Warehouse column at F2
2. On the Summary sheet, create a PIVOTTABLE slicer for the Category field at D2

List all slicers of each type to confirm both were created.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_slicer")
    assert_regex(result.final_response, r"(?i)(warehouse|category|slicer|table|pivot|created)")
    messages = result.messages

    prompt = """
Now let's use both slicers:

1. Use the Table slicer to filter the Inventory table to show only "West" warehouse
2. Use the PivotTable slicer to filter the Summary to show only "Electronics"

How much Electronics stock is in the West warehouse?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_slicer")
    assert_regex(result.final_response, r"(?i)(west|electronics|170|filter|stock)")
    messages = result.messages

    prompt = """
Clear all slicer filters so all data shows again.

Save and close the file.

How many total slicers did we create (both Table and PivotTable)?
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert result.tool_was_called("excel_file")
    assert_regex(result.final_response, r"(?i)(2|two|clear|saved|closed|success)")
