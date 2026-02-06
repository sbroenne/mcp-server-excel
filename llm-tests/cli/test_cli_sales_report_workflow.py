"""CLI complete sales report workflow."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_cli_exit_codes, assert_regex, unique_path

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_sales_report_workflow(aitest_run, excel_cli_server, excel_cli_skill):
    agent = Agent(
        name="cli-sales-report",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        system_prompt=(
            "You are a professional Excel analyst. Execute tasks efficiently using available tools.\n"
            "- Make reasonable assumptions for ambiguous requests\n"
            "- Format data professionally (headers, tables, proper formatting)\n"
            "- Use the Data Model for analysis, not manual calculations\n"
            "- Always verify row counts and data completeness\n"
            "- Report specific numeric values (not just descriptions)"
        ),
    )
    agent.max_turns = 30

    messages = None

    prompt = f"""
I need to analyze our Q1 2025 sales data. Let me build a professional sales analysis workbook.

Step 1 - Create the workbook:
Create a new Excel file at {unique_path('sales-analysis-q1-cli')}

Step 2 - Enter Sales Raw Data (into Sales sheet):
Enter this transaction-level data starting at A1:

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

CRITICAL: To avoid PowerShell JSON parsing issues, you MUST:
- Write the data as a JSON file first (e.g., using Out-File)
- Then use --values-file <path> instead of --values with inline JSON

Step 3 - Create Sales Table and Regional Summary:
- Convert the sales data (A1:H11) into an Excel Table called "SalesTransactions"
- Create a new sheet called Summary
- In the Summary sheet, create a simple table showing: Region (unique values from Sales),
  Transaction Count, Total Revenue (before discount)
- Calculate Total Revenue: Sum of (Quantity * UnitPrice) for all rows

Step 4 - Validate:
- Confirm the SalesTransactions table has exactly 10 data rows (plus header)
- Report the exact total gross revenue amount (sum of Quantity * UnitPrice)
- List all 4 regions found in the data: North, South, East, West
- Verify no calculation errors (row count should match)
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"\$?34[\,.]?200(\.00)?")
    for region in ("North", "South", "East", "West"):
        assert region in result.final_response
    messages = result.messages

    prompt = """
Great! Now let me set up the Data Model for deeper analysis.

IMPORTANT: First, use the excel session list command to discover which Excel file we have open.
Then use that file path for all subsequent operations.

IMPORTANT: Do NOT ask clarifying questions. If the relationship creation fails due to type mismatch,
format both Date columns as Date in the worksheet and retry. If it still fails, proceed without the
relationship and continue with measures and verification.

Step 1 - Add to Data Model:
- Add the SalesTransactions table to the Data Model
- Create a Date dimension table in a new sheet called "DimDate" with 20 unique dates.
  Use this exact list (Date column, A2:A21):
  2025-01-05, 2025-01-06, 2025-01-08, 2025-01-12, 2025-01-15,
  2025-01-18, 2025-01-22, 2025-01-25, 2025-02-01, 2025-02-05,
  2025-01-07, 2025-01-09, 2025-01-10, 2025-01-11, 2025-01-13,
  2025-01-14, 2025-01-16, 2025-01-17, 2025-01-19, 2025-01-20
- Add DimDate to the Data Model
- Create a relationship: SalesTransactions[Date] → DimDate[Date]
- Verify the relationship was created successfully
- Do not loop on reading the Sales date column; set DimDate once and proceed

Step 2 - Create Measures (in SalesTransactions):
Create these DAX measures exactly as specified:
- Revenue (Gross) = SUM(SalesTransactions[Quantity] * SalesTransactions[UnitPrice])
- Discount Amount = SUM(SalesTransactions[Quantity] * SalesTransactions[UnitPrice] * SalesTransactions[Discount])
- Revenue (Net) = [Revenue (Gross)] - [Discount Amount]
- Unit Total = SUM(SalesTransactions[Quantity])
- Average Order Value = DIVIDE([Revenue (Net)], COUNTROWS(SalesTransactions), 0)

Step 3 - Verify Measure Values:
- Query the Data Model to get exact values for each measure:
  * Revenue (Gross) should be exactly $34,200.00
  * Discount Amount should be exactly $1,930.00
  * Revenue (Net) should be exactly $32,270.00
  * Unit Total should be exactly 231 units
  * Average Order Value should be $3,227.00 (or $3,227 rounded)
- Report each measure value with full precision
- If a DAX query fails, compute the value directly from the SalesTransactions table and still report the exact numbers.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"\$?34[\,.]?200(\.00)?")
    assert_regex(result.final_response, r"\$?1[\,.]?930(\.00)?")
    assert_regex(result.final_response, r"\$?32[\,.]?270(\.00)?")
    assert_regex(result.final_response, r"\b231\b")
    messages = result.messages

    prompt = """
Perfect! Now let me create analysis views.

IMPORTANT: First, use the excel session list command to discover which Excel file we have open.
Then use that file path for all subsequent operations.

IMPORTANT: Before creating PivotTables, use excel table list or read to confirm the SalesTransactions table exists.

Create two PivotTables from the SalesTransactions table:

PivotTable 1 - "Revenue by Region":
- Destination: New sheet called "AnalysisRegion"
- Row Fields: Region (top level), Product (nested)
- Data Fields: Sum of Quantity, Sum of Revenue (Gross)
- Format professionally with number formatting for currency

PivotTable 2 - "Salesperson Performance":
- Destination: New sheet called "AnalysisSales"
- Row Fields: Salesperson
- Data Fields: Sum of Quantity, Sum of Revenue (Net), Count of TransactionID
- Apply conditional formatting to Revenue (Net) column, green for highest values
- Sort by Revenue (Net) descending

Step 2 - Analyze Results:
Based on the PivotTables, provide SPECIFIC VALUES for:
- Who is our top salesperson by net revenue? (Provide exact surname and revenue amount)
- Which region has the most transactions? (Provide region name and exact count)
- What's the total quantity sold across all regions? (Provide exact number: 231)
- Rank all 4 salespeople by revenue (highest to lowest with specific amounts)
- Rank all 4 regions by revenue (highest to lowest with specific amounts)

IMPORTANT: If PivotTables do not expose net revenue, run a DAX query against the Data Model to get net revenue by salesperson
(e.g., EVALUATE ADDCOLUMNS(SUMMARIZE(SalesTransactions, SalesTransactions[Salesperson]), "NetRevenue", [Revenue (Net)]))
and use those numbers in your report. Ensure you explicitly state Alice's net revenue as $11,030.00.

Important: Provide specific numeric values, not just descriptions.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert "Alice" in result.final_response
    assert_regex(result.final_response, r"\$?11[\,.]?030(\.00)?")
    assert_regex(result.final_response, r"\b231\b")
    messages = result.messages

    prompt = """
We just received additional February data!

IMPORTANT: First, use the excel session list command to discover which Excel file we have open.
Then use that file path for all subsequent operations.

Add these three new transactions to the SalesTransactions table:

T011, 2025-02-10, East, Laptop Pro, Carol, 4, 1200, 0.05
T012, 2025-02-15, South, Keyboard Mechanical, Bob, 15, 120, 0
T013, 2025-02-20, West, Monitor 4K, Dave, 6, 450, 0.1

Then:
- Refresh all PivotTables to include the new data
- Verify the SalesTransactions table now has EXACTLY 13 rows (10 original + 3 new)
- Query the Data Model to get the UPDATED revenue figures:
  * New Gross Revenue should be $43,500.00 (up from $34,200.00)
  * New Discount should be $2,440.00 (up from $1,930.00)
  * New Net Revenue should be $41,060.00 (up from $32,270.00)
- Check if any salesperson rankings changed
- Verify PivotTables use the new total (13 transactions, not 10)

Important: Do NOT delete and recreate the tables. Use targeted inserts and refreshes only.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(13 rows|13 transactions|13)")
    assert_regex(result.final_response, r"\$?43[\,.]?500(\.00)?")
    assert_regex(result.final_response, r"\$?2[\,.]?440(\.00)?")
    assert_regex(result.final_response, r"\$?41[\,.]?060(\.00)?")
    messages = result.messages

    prompt = """
Perfect! Let me do a final comprehensive check and save our report.

IMPORTANT: First, use the excel session list command to discover which Excel file we have open.
Then use that file path for all subsequent operations.

Step 1 - Structure Verification:
1. List all sheets in the workbook (should be 5): Sales, Summary, DimDate, AnalysisRegion, AnalysisSales
2. Verify the SalesTransactions table has EXACTLY 13 data rows (plus header = 14 rows total)
3. Verify the DimDate table has 20 unique dates
4. Confirm no formulas exist in raw data columns (Quantity, UnitPrice, Discount should be values only)
   - Use excel range get-formulas on Sales!B2:D14 to verify these columns contain values only.
5. Confirm formulas/measures exist IN the Data Model

IMPORTANT: In your response, explicitly include the phrase "13 rows" when reporting the SalesTransactions row count.

Step 2 - Final Revenue Report:
Get the absolute final numbers (use a DAX query or compute directly from SalesTransactions A2:H14):
- Total Gross Revenue across all 13 transactions: $43,500.00
- Total Discounts: $2,440.00
- Total Net Revenue: $41,060.00
- Total Units Sold: 256 (231 + 25)
IMPORTANT: If your DAX query returns earlier values (e.g., $34,200), recompute from the table and report the correct totals above.

Step 3 - Salesperson Deep Dive (by Net Revenue):
- Rank all salespersons by Net Revenue (highest to lowest)
- Report exact revenue for each: Alice, Bob, Carol, Dave
- Include transaction count and unit count for each

Step 4 - Region Deep Dive:
- Rank all regions by revenue (highest to lowest)
- Show transaction count and unit count for each region

Then save the workbook to ensure all changes are persisted.

Report your findings in a structured format.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    for sheet in ("Sales", "Summary", "DimDate", "AnalysisRegion", "AnalysisSales"):
        assert sheet in result.final_response
    assert_regex(result.final_response, r"(?i)(13 data rows|13 rows|13 transactions)")
    assert_regex(result.final_response, r"\$?43[\,.]?500(\.00)?")
    assert_regex(result.final_response, r"\$?2[\,.]?440(\.00)?")
    assert_regex(result.final_response, r"\$?41[\,.]?060(\.00)?")
    assert_regex(result.final_response, r"\b256\b")
    for name in ("Alice", "Bob", "Carol", "Dave"):
        assert name in result.final_response
