using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Regression tests for PivotTable creation with sheet names containing spaces.
/// Same root cause as Bug 1 from Bug Report 2026-02-23: sourceDataRef is constructed
/// as $"{sourceSheet}!{sourceRange}" in PivotTableCommands.Create.cs line 45 without
/// quoting the sheet name. Excel COM requires single quotes around sheet names with
/// spaces: "'Sheet Name'!A1:D6".
/// </summary>
public partial class PivotTableCommandsTests
{
    /// <summary>
    /// Regression: CreateFromRange fails when source sheet name contains spaces.
    /// The source data reference "$sourceSheet!$sourceRange" is not quoted.
    /// Expected: PivotCache.Create succeeds with "'Sales Data'!A1:D6".
    /// Actual: COM error because "Sales Data!A1:D6" is invalid.
    /// </summary>
    [Fact]
    public void CreateFromRange_SourceSheetWithSpaces_CreatesPivotTable()
    {
        // Arrange — create file with data on a sheet whose name has spaces
        var testFile = CreateTestFileWithData_SheetWithSpaces(
            nameof(CreateFromRange_SourceSheetWithSpaces_CreatesPivotTable));

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _pivotCommands.CreateFromRange(
            batch,
            "Sales Data", "A1:D6",    // source sheet with space in name
            "Sales Data", "F1",        // destination on same sheet
            "SpacePivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("SpacePivot", result.PivotTableName);
        Assert.Equal(4, result.AvailableFields.Count);
    }

    /// <summary>
    /// Regression: CreateFromRange fails when destination sheet name contains spaces.
    /// While the current code may handle the destination correctly (it uses
    /// Worksheets[name] not a string reference), this test validates the full path.
    /// </summary>
    [Fact]
    public void CreateFromRange_DestinationSheetWithSpaces_CreatesPivotTable()
    {
        // Arrange
        var testFile = CreateTestFileWithData_SheetWithSpaces(
            nameof(CreateFromRange_DestinationSheetWithSpaces_CreatesPivotTable));

        // Create a second sheet for the PivotTable destination
        using (var setupBatch = ExcelSession.BeginBatch(testFile))
        {
            setupBatch.Execute((ctx, ct) =>
            {
                dynamic? sheets = null;
                dynamic? newSheet = null;
                try
                {
                    sheets = ctx.Book.Worksheets;
                    newSheet = sheets.Add();
                    newSheet.Name = "Pivot Output";
                    return 0;
                }
                finally
                {
                    ComUtilities.Release(ref newSheet);
                    ComUtilities.Release(ref sheets);
                }
            });
            setupBatch.Save();
        } // setupBatch disposed — file lock released

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _pivotCommands.CreateFromRange(
            batch,
            "Sales Data", "A1:D6",     // source sheet with space
            "Pivot Output", "A1",       // destination sheet with space
            "CrossSheetPivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("CrossSheetPivot", result.PivotTableName);
    }

    /// <summary>
    /// Regression: CreateFromRange fails when source sheet name contains special characters.
    /// Excel requires single quotes for sheet names with spaces, hyphens, or other specials.
    /// </summary>
    [Fact]
    public void CreateFromRange_SourceSheetWithHyphen_CreatesPivotTable()
    {
        // Arrange — sheet name with hyphen (also requires quoting)
        var testFile = _fixture.CreateTestFile(
            nameof(CreateFromRange_SourceSheetWithHyphen_CreatesPivotTable));

        using (var setupBatch = ExcelSession.BeginBatch(testFile))
        {
            setupBatch.Execute((ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets[1];
                sheet.Name = "Q1-Sales";

                sheet.Range["A1"].Value2 = "Region";
                sheet.Range["B1"].Value2 = "Product";
                sheet.Range["C1"].Value2 = "Sales";
                sheet.Range["D1"].Value2 = "Date";

                sheet.Range["A2"].Value2 = "North";
                sheet.Range["B2"].Value2 = "Widget";
                sheet.Range["C2"].Value2 = 100;
                sheet.Range["D2"].Value2 = new DateTime(2025, 1, 15);

                sheet.Range["A3"].Value2 = "South";
                sheet.Range["B3"].Value2 = "Gadget";
                sheet.Range["C3"].Value2 = 200;
                sheet.Range["D3"].Value2 = new DateTime(2025, 2, 10);

                return 0;
            });
            setupBatch.Save();
        } // setupBatch disposed — file lock released

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _pivotCommands.CreateFromRange(
            batch,
            "Q1-Sales", "A1:D3",
            "Q1-Sales", "F1",
            "HyphenPivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("HyphenPivot", result.PivotTableName);
    }

    /// <summary>
    /// Helper: Creates a test file with sales data on a sheet named "Sales Data" (with space).
    /// </summary>
    private string CreateTestFileWithData_SheetWithSpaces(string testName)
    {
        var testFile = _fixture.CreateTestFile(testName);

        using var batch = ExcelSession.BeginBatch(testFile);
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            sheet.Name = "Sales Data";  // Space in name — triggers the bug

            sheet.Range["A1"].Value2 = "Region";
            sheet.Range["B1"].Value2 = "Product";
            sheet.Range["C1"].Value2 = "Sales";
            sheet.Range["D1"].Value2 = "Date";

            sheet.Range["A2"].Value2 = "North";
            sheet.Range["B2"].Value2 = "Widget";
            sheet.Range["C2"].Value2 = 100;
            sheet.Range["D2"].Value2 = new DateTime(2025, 1, 15);

            sheet.Range["A3"].Value2 = "North";
            sheet.Range["B3"].Value2 = "Widget";
            sheet.Range["C3"].Value2 = 150;
            sheet.Range["D3"].Value2 = new DateTime(2025, 1, 20);

            sheet.Range["A4"].Value2 = "South";
            sheet.Range["B4"].Value2 = "Gadget";
            sheet.Range["C4"].Value2 = 200;
            sheet.Range["D4"].Value2 = new DateTime(2025, 2, 10);

            sheet.Range["A5"].Value2 = "North";
            sheet.Range["B5"].Value2 = "Gadget";
            sheet.Range["C5"].Value2 = 75;
            sheet.Range["D5"].Value2 = new DateTime(2025, 2, 15);

            sheet.Range["A6"].Value2 = "South";
            sheet.Range["B6"].Value2 = "Widget";
            sheet.Range["C6"].Value2 = 125;
            sheet.Range["D6"].Value2 = new DateTime(2025, 3, 5);

            sheet.Range["D2:D6"].NumberFormat = "m/d/yyyy";

            return 0;
        });

        batch.Save();
        return testFile;
    }
}
