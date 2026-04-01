"""A/B test: does the screenshot tool improve dashboard layout behavior."""

from __future__ import annotations

import pytest

from conftest import build_excel_mcp_eval, assert_regex, unique_path, DEFAULT_TIMEOUT_S

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.mcp]

BASE_TOOLS = ["file", "worksheet", "range", "range_edit", "table", "chart", "chart_config"]


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("name", "allowed_tools", "expects_screenshot"),
    [
        ("dashboard-without-screenshot", BASE_TOOLS, False),
        ("dashboard-with-screenshot", BASE_TOOLS + ["screenshot"], True),
    ],
)
async def test_mcp_dashboard_layout_variants(
    copilot_eval,
    excel_mcp_servers,
    excel_mcp_skill_dir,
    name,
    allowed_tools,
    expects_screenshot,
):
    agent = build_excel_mcp_eval(
        name,
        servers=excel_mcp_servers,
        skill_dir=excel_mcp_skill_dir,
        allowed_tools=allowed_tools,
        max_turns=30,
        timeout_s=DEFAULT_TIMEOUT_S * 3,
    )

    prompt = f"""
Create a new Excel workbook at {unique_path('dashboard-ab-mcp')}.

Set up two data tables:

SalesData at A1 with:
Region, Q1, Q2, Q3
North, 45000, 52000, 48000
South, 38000, 41000, 44000
East, 51000, 49000, 53000
West, 42000, 47000, 50000
Central, 35000, 39000, 42000
Midwest, 40000, 43000, 46000

ExpenseData at F1 with:
Department, Budget, Actual, Variance
Marketing, 120000, 115000, 5000
Engineering, 250000, 262000, -12000
Sales, 180000, 175000, 5000
Operations, 95000, 102000, -7000

Create a dashboard with four charts:
1. Clustered column chart from SalesData below the tables.
2. Pie chart for Q3 sales distribution to the right of the first chart.
3. Line chart for North and South trends below the first chart.
4. Bar chart from ExpenseData to the right of the line chart.

Requirements:
- no chart may overlap any table or another chart,
- every chart must have a descriptive title,
- the workbook should look like a professional dashboard,
- if screenshot tooling is available, use it to verify the final layout before saving.

Save the workbook and report:
- whether the layout has any overlaps,
- whether screenshot-based verification was used,
- confirmation that four charts were created.
"""

    result = await copilot_eval(agent, prompt)
    assert result.success
    assert result.tool_was_called("chart")
    assert result.tool_was_called("table")
    assert_regex(result.final_response, r"(?i)(4 charts|four charts|dashboard)")

    if expects_screenshot:
        assert result.tool_was_called("screenshot")
        assert_regex(result.final_response, r"(?i)(screenshot|visual verification)")
    else:
        assert not result.tool_was_called("screenshot")
