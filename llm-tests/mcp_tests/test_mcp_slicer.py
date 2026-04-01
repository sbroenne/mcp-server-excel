"""MCP slicer workflows for PivotTables and Tables."""

from __future__ import annotations

import pytest

from conftest import build_excel_mcp_eval, assert_regex, unique_results_path

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_pivottable_slicer_workflow(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    agent = build_excel_mcp_eval(
        "mcp-pivot-slicer",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        allowed_tools=["pivottable", "slicer", "table", "range", "file", "worksheet"],
        max_turns=25,
    )

    prompt = f"""
Create a new Excel workbook at {unique_results_path('pivottable-slicer-mcp')}.

On Sheet1, enter:
Region, Product, Quarter, Sales
North, Laptop, Q1, 15000
North, Phone, Q1, 8000
North, Laptop, Q2, 18000
North, Phone, Q2, 9500
South, Laptop, Q1, 12000
South, Phone, Q1, 7500
South, Laptop, Q2, 14000
South, Phone, Q2, 8200

Convert the range to a table named SalesData.
Create a PivotTable on a new sheet named Analysis with Region as rows and Sum of Sales as values.
Create a Region slicer for that PivotTable and position it at E2.
Use the slicer to show only North.
Create a second slicer for Product at G2, then clear the Region filter.
Delete both slicers.
Save the workbook.

Report:
- that the PivotTable slicers were created,
- that North sales equal 50,500,
- that two slicers existed before deletion,
- that both slicers were removed before saving.
"""

    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("pivottable")
    assert result.tool_was_called("slicer")
    assert_regex(result.final_response, r"(?i)(north)")
    assert_regex(result.final_response, r"50[\,.]?500")
    assert_regex(result.final_response, r"(?i)(two slicers|2 slicers|removed)")


@pytest.mark.asyncio
async def test_mcp_table_slicer_workflow(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    agent = build_excel_mcp_eval(
        "mcp-table-slicer",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        allowed_tools=["slicer", "table", "range", "file", "worksheet"],
        max_turns=25,
    )

    prompt = f"""
Create a new Excel workbook at {unique_results_path('table-slicer-mcp')}.

On Sheet1, enter:
Department, Employee, Status, Salary
Engineering, Alice, Active, 85000
Engineering, Bob, Active, 92000
Marketing, Carol, Active, 78000
Marketing, Dave, Inactive, 70000
Sales, Eve, Active, 65000
Sales, Frank, Inactive, 62000
Engineering, Grace, Active, 88000
Sales, Henry, Active, 71000

Convert the range to a table named Employees.
Create a Department table slicer at F2 and use it to filter to Engineering only.
Create a Status slicer at H2 and filter to Active employees.
Delete both slicers.
Save the workbook.

Report:
- that the table slicers were created,
- that Engineering has 3 employees,
- that there are 6 active employees total,
- a short explanation of how table slicers differ from PivotTable slicers,
- confirmation that both slicers were removed before saving.
"""

    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("table")
    assert result.tool_was_called("slicer")
    assert_regex(result.final_response, r"(?i)(engineering)")
    assert_regex(result.final_response, r"(?i)(3|three)")
    assert_regex(result.final_response, r"(?i)(6|six|active)")
    assert_regex(result.final_response, r"(?i)(pivottable slicer|table slicer|removed)")


@pytest.mark.asyncio
async def test_mcp_combined_slicer_workflow(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    agent = build_excel_mcp_eval(
        "mcp-combined-slicer",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        allowed_tools=["pivottable", "slicer", "table", "range", "file", "worksheet"],
        max_turns=25,
    )

    prompt = f"""
Create a new Excel workbook at {unique_results_path('combined-slicer-mcp')}.

On Sheet1, enter:
Category, Product, Warehouse, Stock, Price
Electronics, Laptop, West, 50, 999
Electronics, Phone, West, 120, 599
Electronics, Laptop, East, 35, 999
Electronics, Phone, East, 80, 599
Furniture, Desk, West, 25, 350
Furniture, Chair, West, 40, 175
Furniture, Desk, East, 30, 350
Furniture, Chair, East, 55, 175

Convert the range to a table named Inventory.
Create a PivotTable on a new sheet named Summary with Category as rows and Sum of Stock as values.

Create:
- a Table slicer for Warehouse on Sheet1 at F2,
- a PivotTable slicer for Category on Summary at D2.

Use the Warehouse slicer to filter the Inventory table to West only.
Use the Category slicer to filter the PivotTable to Electronics only.
Then clear all slicer filters.
Save the workbook.

Report:
- that one table slicer and one PivotTable slicer were created,
- that West Electronics stock is 170,
- that all slicer filters were cleared before saving.
"""

    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("table")
    assert result.tool_was_called("pivottable")
    assert result.tool_was_called("slicer")
    assert_regex(result.final_response, r"(?i)(table slicer|pivottable slicer)")
    assert_regex(result.final_response, r"\b170\b")
    assert_regex(result.final_response, r"(?i)(cleared|clear)")
