"""MCP monthly financial report automation workflow."""

from __future__ import annotations

import pytest

from conftest import build_excel_mcp_eval, assert_regex, unique_path

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_financial_report_automation(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    agent = build_excel_mcp_eval(
        "mcp-financial-report",
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        max_turns=30,
    )

    prompt = f"""
Create a new Excel workbook at {unique_path('financial-report-jan2025-mcp')}.

Build a complete monthly financial report on Sheet1.

Section 1 - Revenue summary in A1:B6:
- A1: Revenue Summary
- A2:B4:
  Product Sales, 450000
  Service Revenue, 125000
  Other Income, 18500
- A5:B5: Total Revenue with formula =SUM(B2:B4)

Section 2 - Expense summary in A8:B12:
- A8: Operating Expenses
- A9:B11:
  Salaries, 280000
  Rent, 35000
  Utilities, 12000
- A12:B12: Total Expenses with formula =SUM(B9:B11)

Section 3 - Net income in A14:B15:
- A14: Net Income
- B15 should calculate =B5-B12

Format the report professionally:
- make A1, A8, and A14 bold,
- format all monetary cells as currency with 2 decimals,
- apply alternating row colors to the expense section,
- set column widths to A=25 and B=15.

Then add a variance analysis:
- Put Budget in D1 and Variance in E1.
- Add budget values for Product Sales 440000, Service Revenue 110000, Other Income 20000,
  Salaries 290000, Rent 35000, Utilities 15000.
- Add variance formulas as Actual minus Budget for each line item and for the totals.
- Update Product Sales actuals from 450000 to 455000 and let formulas recalculate.

Finally, add an executive summary table in A17:B22:
- KPI, Value
- Total Revenue (linked to the revenue total),
- Total Expenses (linked to the expense total),
- Net Income (linked to the net income),
- Profit Margin % with a formula based on Net Income / Total Revenue,
- YoY Growth % with a static value of 8.5%.
- Format the summary header with a light blue background, currency cells as currency,
  and percentages with 1 decimal place.

Save the workbook and report:
- Total Revenue,
- Product Sales variance,
- Net Income,
- Profit Margin percentage,
- confirmation that formulas recalculated correctly,
- confirmation that the file was saved.
"""

    result = await copilot_eval(agent, prompt)
    assert result.success
    assert_regex(result.final_response, r"\$?598[\,.]?500(\.00)?")
    assert_regex(result.final_response, r"\$?15[\,.]?000(\.00)?")
    assert_regex(result.final_response, r"\$?271[\,.]?500(\.00)?")
    assert_regex(result.final_response, r"(?i)(45\.4|45\.3|profit margin)")
    assert_regex(result.final_response, r"(?i)(recalculated|formula|saved)")
