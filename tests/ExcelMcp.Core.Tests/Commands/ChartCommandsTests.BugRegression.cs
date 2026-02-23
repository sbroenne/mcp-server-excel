using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Regression tests for Bug Report 2026-02-23.
/// Bug 1: chart(action: 'create-from-range') fails with COM error 0x800A03EC
/// when data is not at row 1 or sheet name contains spaces.
/// </summary>
public partial class ChartCommandsTests
{
    /// <summary>
    /// Regression: create-from-range fails when data starts at a non-first row (e.g., A9:D14).
    /// The user reported COM error 0x800A03EC when creating a Line chart from A9:D14.
    /// This reproduces the exact scenario from the bug report.
    /// </summary>
    [Fact]
    public void CreateFromRange_DataAtNonFirstRow_CreatesChart()
    {
        // Arrange — isolated file with data only at A9:D14 (rows 1-8 empty)
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Write data at A9:D14 (6 rows: 1 header + 5 data)
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ctx.Book.Worksheets[1];
                sheet.Range["A9:D14"].Value2 = new object[,]
                {
                    { "Quarter", "Revenue", "Cost", "Profit" },
                    { "Q1", 1200, 800, 400 },
                    { "Q2", 1500, 900, 600 },
                    { "Q3", 1800, 1000, 800 },
                    { "Q4", 2100, 1100, 1000 },
                    { "Q5", 2400, 1200, 1200 }
                };
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });

        // Act — create Line chart from A9:D14 (exactly as reported)
        var result = _commands.CreateFromRange(
            batch,
            "Sheet1",
            "A9:D14",
            ChartType.Line,
            50, 50, 400, 300,
            "BugRegression_NonFirstRow");

        // Assert
        Assert.True(result.Success, $"CreateFromRange failed: chart was not created");
        Assert.Equal("BugRegression_NonFirstRow", result.ChartName);
        Assert.Equal(ChartType.Line, result.ChartType);

        // Verify chart actually exists in workbook
        var charts = _commands.List(batch);
        Assert.Contains(charts, c => c.Name == "BugRegression_NonFirstRow");
    }

    /// <summary>
    /// Regression: create-from-range fails when sheet name contains spaces.
    /// The range address is constructed as "{sheetName}!{rangeAddress}" but Excel COM
    /// requires single quotes around sheet names with spaces: "'Sheet Name'!A1:D6".
    /// Without quoting, Application.Range["Deal Summary!A9:D14"] throws 0x800A03EC.
    /// </summary>
    [Fact]
    public void CreateFromRange_SheetNameWithSpaces_CreatesChart()
    {
        // Arrange — create a sheet with spaces in name
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic? sheets = null;
            dynamic? newSheet = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                newSheet = sheets.Add();
                newSheet.Name = "Deal Summary";
                newSheet.Range["A1:D6"].Value2 = new object[,]
                {
                    { "Product", "Q1", "Q2", "Q3" },
                    { "Widget A", 100, 150, 200 },
                    { "Widget B", 200, 250, 300 },
                    { "Widget C", 300, 350, 400 },
                    { "Widget D", 400, 450, 500 },
                    { "Widget E", 500, 550, 600 }
                };
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref newSheet);
                ComUtilities.Release(ref sheets);
            }
        });

        // Act — create chart on sheet with spaces in name
        var result = _commands.CreateFromRange(
            batch,
            "Deal Summary",
            "A1:D6",
            ChartType.Line,
            50, 50, 400, 300,
            "BugRegression_SpacesInName");

        // Assert
        Assert.True(result.Success, $"CreateFromRange failed for sheet with spaces in name");
        Assert.Equal("BugRegression_SpacesInName", result.ChartName);
        Assert.Equal("Deal Summary", result.SheetName);
    }

    /// <summary>
    /// Regression: Combined scenario — sheet name with spaces AND data at non-first row.
    /// This is the exact scenario from the Bayer AG deal sizing bug report.
    /// </summary>
    [Fact]
    public void CreateFromRange_SheetWithSpacesAndDataAtNonFirstRow_CreatesChart()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic? sheets = null;
            dynamic? newSheet = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                newSheet = sheets.Add();
                newSheet.Name = "Export Data";
                // Data starts at row 9, like the bug report
                newSheet.Range["A9:D14"].Value2 = new object[,]
                {
                    { "Service", "Current", "Proposed", "Delta" },
                    { "Compute", 50000, 45000, -5000 },
                    { "Storage", 20000, 18000, -2000 },
                    { "Network", 15000, 14000, -1000 },
                    { "Database", 30000, 25000, -5000 },
                    { "AI/ML", 10000, 12000, 2000 }
                };
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref newSheet);
                ComUtilities.Release(ref sheets);
            }
        });

        // Act
        var result = _commands.CreateFromRange(
            batch,
            "Export Data",
            "A9:D14",
            ChartType.Line,
            50, 50, 400, 300,
            "BugRegression_Combined");

        // Assert
        Assert.True(result.Success, $"CreateFromRange failed for combined scenario");
        Assert.Equal("BugRegression_Combined", result.ChartName);
        Assert.Equal("Export Data", result.SheetName);
        Assert.Equal(ChartType.Line, result.ChartType);
    }

    /// <summary>
    /// Regression: Verify that create-from-table works as workaround for the same data layout.
    /// The bug report confirms create-from-table succeeds where create-from-range fails.
    /// This test validates the workaround and serves as a comparison baseline.
    /// </summary>
    [Fact]
    public void CreateFromTable_DataAtNonFirstRow_SucceedsAsWorkaround()
    {
        // Arrange — same data layout as the failing create-from-range scenario
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? listObjects = null;
            dynamic? table = null;
            try
            {
                sheet = ctx.Book.Worksheets[1];
                sheet.Range["A9:D14"].Value2 = new object[,]
                {
                    { "Quarter", "Revenue", "Cost", "Profit" },
                    { "Q1", 1200, 800, 400 },
                    { "Q2", 1500, 900, 600 },
                    { "Q3", 1800, 1000, 800 },
                    { "Q4", 2100, 1100, 1000 },
                    { "Q5", 2400, 1200, 1200 }
                };
                listObjects = sheet.ListObjects;
                table = listObjects.Add(1, sheet.Range["A9:D14"], null, 1); // xlYes = 1
                table.Name = "BugWorkaroundTable";
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref listObjects);
                ComUtilities.Release(ref sheet);
            }
        });

        // Act — create chart from table (workaround from bug report)
        var result = _commands.CreateFromTable(
            batch,
            "BugWorkaroundTable",
            "Sheet1",
            ChartType.Line,
            50, 50, 400, 300,
            "BugWorkaround_Table");

        // Assert — this should always succeed
        Assert.True(result.Success, $"CreateFromTable (workaround) failed unexpectedly");
        Assert.Equal("BugWorkaround_Table", result.ChartName);
        Assert.Equal(ChartType.Line, result.ChartType);
    }
}
