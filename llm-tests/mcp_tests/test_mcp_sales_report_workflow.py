"""MCP complete sales report workflow."""

from __future__ import annotations

import pytest

from conftest import build_excel_mcp_eval, assert_regex, unique_results_path

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_sales_report_workflow(copilot_eval, excel_mcp_servers, excel_mcp_skill_dir):
    agent = build_excel_mcp_eval(
        "mcp-sales-report",
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
        instructions=(
            "You are a professional Excel analyst. Execute tasks efficiently with the available "
            "tools, keep the workbook well-structured, and report specific numeric values."
        ),
        max_turns=40,
    )

    prompt = f"""
Create a new Excel workbook at {unique_results_path('sales-analysis-q1-mcp')}.

Build a complete sales analysis workbook in one run.

1. On a sheet named Sales, enter these 10 transactions starting at A1:
   TransactionID, Date, Region, Product, Salesperson, Quantity, UnitPrice, Discount
   T001, 2025-01-05, North, Laptop Pro, Alice, 5, 1200, 0.05
   T002, 2025-01-06, North, Mouse Wireless, Alice, 50, 25, 0
   T003, 2025-01-08, South, Laptop Pro, Bob, 3, 1200, 0.1
   T004, 2025-01-12, South, Monitor 4K, Bob, 8, 450, 0.05
   T005, 2025-01-15, East, Keyboard Mechanical, Carol, 30, 120, 0
   T006, 2025-01-18, North, Monitor 4K, Alice, 4, 450, 0
   T007, 2025-01-22, East, Laptop Pro, Carol, 6, 1200, 0.1
   T008, 2025-01-25, West, Mouse Wireless, Dave, 100, 25, 0.1
   T009, 2025-02-01, South, Monitor 4K, Bob, 5, 450, 0
   T010, 2025-02-05, North, Keyboard Mechanical, Alice, 20, 120, 0.05
   Convert the range to a table named SalesTransactions.

2. Create a sheet named Summary with a simple region summary table showing:
   Region, Transaction Count, Gross Revenue.

3. Create a sheet named DimDate with 20 unique dates including all transaction dates and add it to the Data Model.
   Add SalesTransactions to the Data Model and create a relationship SalesTransactions[Date] -> DimDate[Date].

4. Create these DAX measures on SalesTransactions:
   - Revenue (Gross)
   - Discount Amount
   - Revenue (Net)
   - Unit Total
   - Average Order Value

5. Create two PivotTables:
   - AnalysisRegion: Region then Product as rows; Quantity and Revenue (Gross) as values.
   - AnalysisSales: Salesperson as rows; Quantity, Revenue (Net), and Transaction Count as values.

6. Add three new transactions to the SalesTransactions table without recreating it:
   T011, 2025-02-10, East, Laptop Pro, Carol, 4, 1200, 0.05
   T012, 2025-02-15, South, Keyboard Mechanical, Bob, 15, 120, 0
   T013, 2025-02-20, West, Monitor 4K, Dave, 6, 450, 0.1
   Refresh dependent analysis so the new rows are included.

7. Add a chart on AnalysisRegion that visualizes revenue by region.

8. Save the workbook.

Report all of the following explicitly:
- the workbook sheets Sales, Summary, DimDate, AnalysisRegion, and AnalysisSales,
- that SalesTransactions now has 13 rows,
- Gross Revenue = $43,500.00,
- Discount Amount = $2,440.00,
- Revenue (Net) = $41,060.00,
- Unit Total = 256,
- Alice, Bob, Carol, and Dave ranked by revenue,
- confirmation that the PivotTables and chart were refreshed after the new rows were added.
"""

    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("table")
    assert result.tool_was_called("datamodel")
    assert result.tool_was_called("pivottable")
    assert result.tool_was_called("chart")
    for sheet in ("Sales", "Summary", "DimDate", "AnalysisRegion", "AnalysisSales"):
        assert sheet in (result.final_response or "")
    assert_regex(result.final_response, r"(?i)(13 rows|13 transactions)")
    assert_regex(result.final_response, r"\$?43[\,.]?500(\.00)?")
    assert_regex(result.final_response, r"\$?2[\,.]?440(\.00)?")
    assert_regex(result.final_response, r"\$?41[\,.]?060(\.00)?")
    assert_regex(result.final_response, r"\b256\b")
    for name in ("Alice", "Bob", "Carol", "Dave"):
        assert name in (result.final_response or "")
