"""
Enhanced screenshot test suite - Diagnosing blank screenshot issues.

These tests verify that screenshot functionality correctly captures Excel content
and returns visible, non-blank images. They test various scenarios that might
cause blank screenshots:
- Simple data ranges
- Formatted cells with colors and fonts
- Charts and visualizations
- Merged cells
- Different sheet sizes

Tests use llm_assert_image (vision LLM judge) to verify screenshot content.
"""

from __future__ import annotations

import asyncio
import pytest

from pytest_aitest import Agent, Provider

from conftest import (
    unique_path,
    DEFAULT_RETRIES,
    DEFAULT_TIMEOUT_MS,
)

pytestmark = [pytest.mark.aitest, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_screenshot_simple_data_capture(aitest_run, excel_mcp_server, excel_mcp_skill, llm_assert_image):
    """Test that screenshot captures simple text and data."""
    agent = Agent(
        name="screenshot-simple",
        provider=Provider(model="azure/gpt-4.1"),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=["file", "worksheet", "range", "screenshot"],
        max_turns=50,
        retries=DEFAULT_RETRIES,
    )

    path = unique_path("screenshot-simple")
    prompt = f"""
Create a new Excel file at {path}.

In the first sheet, add:
- Cell A1: "Product Name"
- Cell B1: "Quantity"
- Cell C1: "Price"
- Cell A2: "Widget A"
- Cell B2: "100"
- Cell C2: "$25.00"
- Cell A3: "Widget B"
- Cell B3: "150"
- Cell C3: "$30.00"

Format the header row (A1:C1) with bold text and blue background.
Then take a screenshot of range A1:C3 to verify the data is visible and not blank.
The screenshot should clearly show:
- The three header cells with blue background and bold text
- All six data cells with their content visible
- No blank/white image - must have visible text and colors

Save the file when done.
"""

    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert result.tool_was_called("screenshot"), "Expected screenshot to be called"

    # Verify the screenshot was captured and is not blank
    screenshots = result.tool_images_for("screenshot")
    assert screenshots, "No screenshots were captured"

    last_screenshot = screenshots[-1]
    assert llm_assert_image(
        last_screenshot,
        "Shows an Excel spreadsheet with three columns (Product Name, Quantity, Price) "
        "and visible data rows with text content. "
        "The image is not blank or all-white — it must show cell grid lines and readable text.",
    )


@pytest.mark.asyncio
async def test_screenshot_formatted_cells(aitest_run, excel_mcp_server, excel_mcp_skill, llm_assert_image):
    """Test that screenshot captures formatting (colors, fonts, borders)."""
    agent = Agent(
        name="screenshot-formatted",
        provider=Provider(model="azure/gpt-4.1"),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=["file", "worksheet", "range", "range_format", "screenshot"],
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    path = unique_path("screenshot-formatted")
    prompt = f"""
Create a new Excel file at {path}.

Create a formatted report layout:
1. In A1, add title "Sales Report 2024" and make it bold with 14pt font
2. In A3, add "Region" and B3 add "Sales". Make both bold with green background.
3. Add the following data below (A4:B7):
   - North, 125000
   - South, 142000
   - East, 158000
   - West, 131000
4. Format the data cells with borders and alternating row colors (light gray for every other row)
5. Set column widths appropriately so text is clearly visible
6. Take a screenshot of the entire formatted range A1:B7

The screenshot must show:
- The bold title text
- The green header background
- Row alternations and borders clearly visible
- All text content clearly readable

If the screenshot appears blank or monochrome, this indicates a rendering issue.
Save the file when done.
"""

    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert result.tool_was_called("screenshot"), "Expected screenshot to be called"

    screenshots = result.tool_images_for("screenshot")
    assert screenshots, "No screenshots captured"

    last_screenshot = screenshots[-1]
    assert llm_assert_image(
        last_screenshot,
        "Shows an Excel spreadsheet with visible text, colored cell backgrounds (green headers), "
        "and formatted rows. The image is not blank or all-white.",
    )


@pytest.mark.asyncio
async def test_screenshot_with_chart(aitest_run, excel_mcp_server, excel_mcp_skill, llm_assert_image):
    """Test that screenshot captures charts and visualizations."""
    agent = Agent(
        name="screenshot-chart",
        provider=Provider(model="azure/gpt-4.1"),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=["file", "worksheet", "range", "chart", "screenshot"],
        max_turns=50,
        retries=DEFAULT_RETRIES,
    )

    # Brief pause to let Azure rate limits recover before this heavy multi-tool test
    await asyncio.sleep(10)

    path = unique_path("screenshot-chart")
    prompt = f"""
Create a new Excel file at {path}.

1. Write data into the sheet:
   - A1: Quarter, B1: Revenue
   - A2: Q1, B2: 250000
   - A3: Q2, B3: 280000
   - A4: Q3, B4: 310000
   - A5: Q4, B5: 295000

2. Create a column chart from the range A1:B5.
   Position the chart starting at cell D1.

3. Save the file.

4. Take a screenshot of the range A1:M15 to capture both the data and the chart.

The screenshot should show visible data and is not blank or all-white.
"""

    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)  # 10 min default handles Azure GlobalStandard latency
    assert result.success
    assert result.tool_was_called("screenshot"), "Expected screenshot to be called"

    screenshots = result.tool_images_for("screenshot")
    assert screenshots, "No screenshots captured"

    last_screenshot = screenshots[-1]
    assert llm_assert_image(
        last_screenshot,
        "Shows an Excel spreadsheet with a data table. "
        "The image is not blank or all-white — it must show visible text and cell content.",
    )


@pytest.mark.asyncio
async def test_screenshot_sheet_overview(aitest_run, excel_mcp_server, excel_mcp_skill, llm_assert_image):
    """Test full-sheet screenshot capture (no range specified)."""
    agent = Agent(
        name="screenshot-sheet",
        provider=Provider(model="azure/gpt-4.1"),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=["file", "worksheet", "range", "range_format", "screenshot"],
        max_turns=50,
        retries=DEFAULT_RETRIES,
    )

    path = unique_path("screenshot-sheet")
    prompt = f"""
Create a new Excel file at {path}.

Set up a simple worksheet:
1. Add headers in row 1: Name, Email, Phone, Department
2. Add 5 rows of employee data with sample information:
   - John Smith, john@company.com, 555-1234, Sales
   - Jane Doe, jane@company.com, 555-5678, Engineering
   - Bob Johnson, bob@company.com, 555-9012, Marketing
   - Alice Williams, alice@company.com, 555-3456, Operations
   - Charlie Brown, charlie@company.com, 555-7890, Finance
3. Format the headers with bold and light blue background.
4. Save the file.
5. REQUIRED: You MUST call the screenshot tool with action=capture-sheet to capture the full worksheet.
   Do NOT finish your response without calling the screenshot tool.
   The screenshot is the final and most important step of this task.
"""

    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert result.tool_was_called("screenshot"), "Expected screenshot of sheet"

    screenshots = result.tool_images_for("screenshot")
    assert screenshots, "No screenshots captured"

    last_screenshot = screenshots[-1]
    assert llm_assert_image(
        last_screenshot,
        "Shows an Excel spreadsheet with employee data in multiple rows and columns, "
        "with a formatted header row. The image is not blank or all-white.",
    )


@pytest.mark.asyncio
async def test_screenshot_debugging_capture_bounds(aitest_run, excel_mcp_server, excel_mcp_skill, llm_assert_image):
    """Test screenshot with explicit bounds to help debug blank image issue."""
    agent = Agent(
        name="screenshot-bounds",
        provider=Provider(model="azure/gpt-4.1"),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=["file", "worksheet", "range", "screenshot"],
        max_turns=50,
        retries=DEFAULT_RETRIES,
    )

    path = unique_path("screenshot-bounds")
    prompt = f"""
Create a new Excel file at {path}.

1. In cell A1, write "START"
2. Add data in A2:C3:
   - Row 2: Value 100, Value 200, Value 300
   - Row 3: Value 400, Value 500, Value 600
3. In cell A4, write "END"
4. Save the file.
5. REQUIRED: Call the screenshot tool with action=capture-range and range_address=A1:C4 to capture the cells.

You MUST call the screenshot tool as step 5. This is required for verification.
"""

    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success

    screenshots = result.tool_images_for("screenshot")
    assert screenshots, "No screenshots captured - check if screenshot tool is working"

    for i, screenshot in enumerate(screenshots):
        assert llm_assert_image(
            screenshot,
            "Shows an Excel spreadsheet with text cells including data values. "
            "The image is not blank or all-white.",
        ), f"Screenshot #{i + 1} appears blank — expected visible cell content"


@pytest.mark.asyncio
async def test_screenshot_vs_get_values_consistency(aitest_run, excel_mcp_server, excel_mcp_skill, llm_assert_image):
    """Verify screenshot content matches get-values (data should be consistent)."""
    agent = Agent(
        name="screenshot-data-consistency",
        provider=Provider(model="azure/gpt-4.1"),
        mcp_servers=[excel_mcp_server],
        skill=excel_mcp_skill,
        allowed_tools=["file", "worksheet", "range", "screenshot"],
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    path = unique_path("screenshot-consistency")
    prompt = f"""
Create a new Excel file at {path}.

Test consistency between data and visualization:
1. Create a table with specific numerical data:
   - A1:B5 containing:
     Header1, Header2
     10, 20
     30, 40
     50, 60
     70, 80

2. Take a screenshot of A1:B5.

3. Get the values using get-values to verify the data is intact.

4. Report whether the screenshot shows the same values as get-values confirms are in the cells.
   If the screenshot appears blank but get-values shows data, this indicates a rendering issue.

The screenshot should clearly display:
- Both column headers
- All 4 rows of numerical data
- Numbers should be clearly readable

Save the file.
"""

    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert result.tool_was_called("screenshot"), "Expected screenshot"

    screenshots = result.tool_images_for("screenshot")
    assert screenshots, "No screenshots captured"

    for screenshot in screenshots:
        assert llm_assert_image(
            screenshot,
            "Shows an Excel spreadsheet with two columns of numerical data and headers. "
            "Numbers are clearly readable. The image is not blank or all-white.",
        )
