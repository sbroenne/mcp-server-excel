"""A/B test: Does the screenshot tool improve dashboard layout quality?

This test compares two agent variants building the same complex dashboard:
- Control (without-screenshot): Cannot use the screenshot tool
- Experiment (with-screenshot): Can visually verify and self-correct via screenshot

Both agents receive the same natural-language prompt to create a dashboard with
4 charts and 2 data tables with strict no-overlap requirements.

How to interpret results (in TestResults/report.html):
- Compare pass rates: does the screenshot variant succeed more often?
- Compare token usage: how much overhead does visual verification add?
- Check if experiment variant calls screenshot and self-corrects overlaps
- The AI summary model analyzes variant effectiveness automatically
- Image assertions verify the visual output quality (experiment only)
"""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import (
    assert_regex,
    unique_path,
    DEFAULT_RETRIES,
    DEFAULT_TIMEOUT_MS,
)

pytestmark = [pytest.mark.aitest, pytest.mark.mcp]

# Tools available for dashboard creation (no screenshot)
BASE_TOOLS = ["file", "worksheet", "range", "range_edit", "table", "chart", "chart_config"]

AGENTS = [
    Agent(
        name="dashboard-without-screenshot",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        allowed_tools=BASE_TOOLS,
        max_turns=30,
        retries=DEFAULT_RETRIES,
    ),
    Agent(
        name="dashboard-with-screenshot",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        allowed_tools=BASE_TOOLS + ["screenshot"],
        max_turns=30,
        retries=DEFAULT_RETRIES,
    ),
]


@pytest.mark.asyncio
@pytest.mark.xfail(reason="Complex dashboard workflow - LLM may hit turn/retry limits", strict=False)
@pytest.mark.parametrize("agent", AGENTS, ids=lambda a: a.name)
async def test_mcp_dashboard_layout(
    aitest_run, excel_mcp_server, excel_mcp_skill, agent, llm_assert_image
):
    # Attach server and skill (not set at module level since they're fixtures)
    agent = Agent(
        name=agent.name,
        provider=agent.provider,
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=agent.allowed_tools,
        max_turns=agent.max_turns,
        retries=agent.retries,
    )

    path = unique_path("dashboard-ab")
    prompt = f"""
Create a new Excel file at {path}.

Set up two data tables:

Table 1 - "SalesData" starting at A1 with 7 rows (header + 6 data rows):
Region, Q1, Q2, Q3
North, 45000, 52000, 48000
South, 38000, 41000, 44000
East, 51000, 49000, 53000
West, 42000, 47000, 50000
Central, 35000, 39000, 42000
Midwest, 40000, 43000, 46000

Table 2 - "ExpenseData" starting at F1 with 5 rows (header + 4 data rows):
Department, Budget, Actual, Variance
Marketing, 120000, 115000, 5000
Engineering, 250000, 262000, -12000
Sales, 180000, 175000, 5000
Operations, 95000, 102000, -7000

Now create a professional dashboard with these 4 charts. Position them so that
NO chart overlaps any other chart or any data table:

1. A clustered column chart from SalesData showing all regions and quarters.
   Place it below the data tables.
2. A pie chart showing Q3 sales distribution by region.
   Place it to the right of the column chart.
3. A line chart showing the sales trend across Q1-Q3 for North and South regions.
   Place it below the column chart.
4. A bar chart from ExpenseData showing Budget vs Actual by department.
   Place it to the right of the line chart.

Requirements:
- No chart may overlap any other chart or any data table
- Each chart must have a descriptive title
- The dashboard should look professional and well-organized

After placing all charts, take a screenshot to verify the final layout is correct
with no overlaps. Save the file.
"""

    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS * 3)
    assert result.success
    assert result.tool_was_called("chart")
    assert result.tool_was_called("table")
    assert_regex(result.final_response, r"(?i)(chart|dashboard|layout)")

    # Experiment variant: verify visual quality with screenshot tool
    if agent.name == "dashboard-with-screenshot":
        assert result.tool_was_called("screenshot"), (
            "Expected experiment variant to use screenshot for visual verification"
        )

        # AI-graded visual evaluation of the final screenshot
        screenshots = result.tool_images_for("screenshot")
        if screenshots:
            assert llm_assert_image(
                screenshots[-1],
                "Shows a well-organized dashboard with 4 charts. "
                "Charts should not overlap each other or the data tables. "
                "Each chart should have a descriptive title.",
            )
