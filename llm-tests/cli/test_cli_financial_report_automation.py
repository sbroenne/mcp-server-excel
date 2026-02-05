"""CLI monthly financial report automation workflow."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_cli_exit_codes, assert_regex, unique_path

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_financial_report_automation(aitest_run, excel_cli_server, excel_cli_skill):
    agent = Agent(
        name="cli-financial-report",
        provider=Provider(model="azure/gpt-5-mini", rpm=10, tpm=10000),
        cli_servers=[excel_cli_server],
        skill=excel_cli_skill,
        system_prompt=(
            "You are an automation engineer using excelcli. Write efficient, idempotent CLI scripts.\n"
            "- Always use -q flag for clean JSON output\n"
            "- Use --save flag to persist changes\n"
            "- Keep sessions open while working on same file\n"
            "- Validate data after each operation\n"
            "- Provide clear status updates"
        ),
    )

    messages = None

    prompt = f"""
I need to automate the creation of our monthly financial report. Let me build this from scratch with data, formulas, and formatting.

STEP 1: Create the report file and structure

Create {unique_path('financial-report-jan2025')}
Open it and create the following structure in Sheet1:

In A1:B6, create the revenue section:
Revenue Summary,
Product Sales, 450000
Service Revenue, 125000
Other Income, 18500
Total Revenue, (formula: =SUM(B2:B4))

In A8:B12, create the expense section:
Operating Expenses,
Salaries, 280000
Rent, 35000
Utilities, 12000
Total Expenses, (formula: =SUM(B9:B11))

In A14:B15:
Net Income, (formula: =B6-B12)

STEP 2: Format the report professionally

- Make headers (A1, A8, A14) bold
- Format all currency amounts (B2:B4, B9:B11, B6, B12, B15) as currency (2 decimals)
- Apply alternating row colors to the expense section
- Set column widths: A=25, B=15

STEP 3: Verify calculations

Read the calculated values:
- Total Revenue
- Total Expenses
- Net Income

Report all three values to confirm formulas are working.
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(450000|revenue|net income|formula)")
    messages = result.messages

    prompt = """
Great! Now add a variance analysis. I need to compare actual vs. budget.

STEP 1: Add Budget data

In column D, add Budget column header, then budget figures:
- Product Sales Budget: 440000
- Service Revenue Budget: 110000
- Other Income Budget: 20000
- Salaries Budget: 290000
- Rent Budget: 35000
- Utilities Budget: 15000

STEP 2: Add Variance formulas

In column E (labeled "Variance"), add formulas that calculate:
- Variance = Actual - Budget for each line item
- For totals, show the variance calculations too

STEP 3: Update one actual value

Product Sales actually came in at 455000 (not 450000). Update B2 to reflect this.
Verify that:
- The variance in E2 automatically recalculates
- The Total Revenue in B6 updates automatically
- The Net Income in B15 updates automatically

STEP 4: Report final numbers

Provide:
- New Total Revenue
- Product Sales Variance (actual vs. budget)
- New Net Income
- Confirm all formulas are still working
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(variance|455000|formula|recalculate|updated)")
    messages = result.messages

    prompt = """
Perfect! Let me create a summary dashboard at the bottom and finalize the report.

STEP 1: Create Executive Summary Table

In A17:B22, create a summary table:
KPI, Value
Total Revenue, (formula referencing B6)
Total Expenses, (formula referencing B12)
Net Income, (formula referencing B15)
Profit Margin %, (formula: =B20/B18*100, formatted as percentage)
YoY Growth %, 8.5%

STEP 2: Format the summary
- Make the table header (A17:B17) bold with light blue background
- Format currency cells appropriately
- Format percentage to 1 decimal place

STEP 3: Final validation

Read the used range of the sheet - should show everything is there and formatted.

Get the Profit Margin % value to verify it calculated correctly:
Should be approximately: (573500 / 455000 + 125000 + 18500) * 100

STEP 4: Save the report

Save the file with all changes.

Report:
- Total Revenue amount
- Net Income amount
- Profit Margin percentage
- Confirmation file was saved
"""
    result = await aitest_run(agent, prompt, messages=messages)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(summary|margin|saved|598500|complete)")
