"""CLI slicer workflows."""

from __future__ import annotations

import pytest

from conftest import build_excel_cli_eval, assert_cli_exit_codes, assert_regex, unique_path

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_pivottable_slicer_workflow(copilot_eval, excel_cli_servers, excel_cli_skill_dir):
    agent = build_excel_cli_eval(
        "cli-pivot-slicer",
        servers=excel_cli_servers,
        skill_dir=excel_cli_skill_dir,
        max_turns=25,
    )

    prompt = f"""
Using the Excel CLI tool, create a new workbook at {unique_path('pivottable-slicer-cli')}.

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
Create a Region slicer for that PivotTable at E2.
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
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(north)")
    assert_regex(result.final_response, r"50[\,.]?500")
    assert_regex(result.final_response, r"(?i)(two slicers|2 slicers|removed)")


@pytest.mark.asyncio
async def test_cli_table_slicer_workflow(copilot_eval, excel_cli_servers, excel_cli_skill_dir):
    agent = build_excel_cli_eval(
        "cli-table-slicer",
        servers=excel_cli_servers,
        skill_dir=excel_cli_skill_dir,
        max_turns=25,
    )

    prompt = f"""
Using the Excel CLI tool, create a new workbook at {unique_path('table-slicer-cli')}.

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

Convert the range to an Excel table named Employees.
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
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(engineering)")
    assert_regex(result.final_response, r"(?i)(3|three)")
    assert_regex(result.final_response, r"(?i)(6|six|active)")
    assert_regex(result.final_response, r"(?i)(pivottable slicer|table slicer|removed)")
