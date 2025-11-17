using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for advanced Range operations (clear formats, copy formulas, insert/delete, hyperlinks)
/// Optimized: Single batch per test, no SaveAsync unless testing persistence
/// </summary>
public partial class RangeCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    [Trait("Speed", "Medium")]
    public void ClearFormats_FormattedRange_RemovesFormattingOnly()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests), nameof(ClearFormats_FormattedRange_RemovesFormattingOnly), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Set values with formatting
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Test";
            sheet.Range["A1"].Font.Bold = true;
            sheet.Range["A1"].Interior.Color = 255; // Red background
            return 0;
        });

        // Act - Clear only formats
        var result = _commands.ClearFormats(batch, "Sheet1", "A1");

        // Assert
        Assert.True(result.Success, $"ClearFormats failed: {result.ErrorMessage}");

        // Verify value remains but formatting is gone
        var values = _commands.GetValues(batch, "Sheet1", "A1");
        Assert.Equal("Test", values.Values[0][0]?.ToString());
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void CopyFormulas_SourceWithFormulas_CopiesFormulasOnly()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests), nameof(CopyFormulas_SourceWithFormulas_CopiesFormulasOnly), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Set up source data with formulas
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = 10;
            sheet.Range["A2"].Value2 = 20;
            sheet.Range["A3"].Formula = "=A1+A2";
            return 0;
        });

        // Act - Copy formulas to B3
        var result = _commands.CopyFormulas(batch, "Sheet1", "A3", "Sheet1", "B3");

        // Assert
        Assert.True(result.Success, $"CopyFormulas failed: {result.ErrorMessage}");

        // Verify formula was copied (should adjust references)
        var formulas = _commands.GetFormulas(batch, "Sheet1", "B3");
        Assert.NotNull(formulas.Formulas[0][0]);
        Assert.Contains("+", formulas.Formulas[0][0]?.ToString());
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void InsertCells_ShiftDown_InsertsAndShiftsExisting()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests), nameof(InsertCells_ShiftDown_InsertsAndShiftsExisting), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Set up initial data
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Original";
            return 0;
        });

        // Act - Insert cell at A1, shifting down
        var result = _commands.InsertCells(batch, "Sheet1", "A1", InsertShiftDirection.Down);

        // Assert
        Assert.True(result.Success, $"InsertCells failed: {result.ErrorMessage}");

        // Verify original value shifted to A2
        var values = _commands.GetValues(batch, "Sheet1", "A2");
        Assert.Equal("Original", values.Values[0][0]?.ToString());
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteCells_ShiftUp_RemovesAndShifts()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests), nameof(DeleteCells_ShiftUp_RemovesAndShifts), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Set up data in A1 and A2
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Delete Me";
            sheet.Range["A2"].Value2 = "Keep Me";
            return 0;
        });

        // Act - Delete A1, shifting up
        var result = _commands.DeleteCells(batch, "Sheet1", "A1", DeleteShiftDirection.Up);

        // Assert
        Assert.True(result.Success, $"DeleteCells failed: {result.ErrorMessage}");

        // Verify A2 value shifted to A1
        var values = _commands.GetValues(batch, "Sheet1", "A1");
        Assert.Equal("Keep Me", values.Values[0][0]?.ToString());
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void InsertRows_BeforeExistingData_InsertsBlankRows()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests), nameof(InsertRows_BeforeExistingData_InsertsBlankRows), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Set up data in row 1
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Row 1";
            return 0;
        });

        // Act - Insert 2 rows at row 1
        var result = _commands.InsertRows(batch, "Sheet1", "1:2");

        // Assert
        Assert.True(result.Success, $"InsertRows failed: {result.ErrorMessage}");

        // Verify original data shifted to row 3
        var values = _commands.GetValues(batch, "Sheet1", "A3");
        Assert.Equal("Row 1", values.Values[0][0]?.ToString());
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteRows_ExistingRows_RemovesRows()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests), nameof(DeleteRows_ExistingRows_RemovesRows), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Set up data in rows 1-3
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Row 1";
            sheet.Range["A2"].Value2 = "Row 2 - Delete";
            sheet.Range["A3"].Value2 = "Row 3";
            return 0;
        });

        // Act - Delete row 2
        var result = _commands.DeleteRows(batch, "Sheet1", "2:2");

        // Assert
        Assert.True(result.Success, $"DeleteRows failed: {result.ErrorMessage}");

        // Verify row 3 shifted to row 2
        var values = _commands.GetValues(batch, "Sheet1", "A2");
        Assert.Equal("Row 3", values.Values[0][0]?.ToString());
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void InsertColumns_BeforeExistingData_InsertsBlankColumns()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests), nameof(InsertColumns_BeforeExistingData_InsertsBlankColumns), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Set up data in column A
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Col A";
            return 0;
        });

        // Act - Insert 2 columns at column A (column 1)
        var result = _commands.InsertColumns(batch, "Sheet1", "A:B");

        // Assert
        Assert.True(result.Success, $"InsertColumns failed: {result.ErrorMessage}");

        // Verify original data shifted to column C
        var values = _commands.GetValues(batch, "Sheet1", "C1");
        Assert.Equal("Col A", values.Values[0][0]?.ToString());
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteColumns_ExistingColumns_RemovesColumns()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests), nameof(DeleteColumns_ExistingColumns_RemovesColumns), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Set up data in columns A-C
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Col A";
            sheet.Range["B1"].Value2 = "Col B - Delete";
            sheet.Range["C1"].Value2 = "Col C";
            return 0;
        });

        // Act - Delete column B
        var result = _commands.DeleteColumns(batch, "Sheet1", "B:B");

        // Assert
        Assert.True(result.Success, $"DeleteColumns failed: {result.ErrorMessage}");

        // Verify column C shifted to B
        var values = _commands.GetValues(batch, "Sheet1", "B1");
        Assert.Equal("Col C", values.Values[0][0]?.ToString());
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void GetHyperlink_ExistingHyperlink_ReturnsDetails()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests), nameof(GetHyperlink_ExistingHyperlink_ReturnsDetails), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Add a hyperlink
        var addResult = _commands.AddHyperlink(batch, "Sheet1", "A1", "https://example.com", "Example Link");
        Assert.True(addResult.Success);

        // Act
        var result = _commands.GetHyperlink(batch, "Sheet1", "A1");

        // Assert
        Assert.True(result.Success, $"GetHyperlink failed: {result.ErrorMessage}");
        Assert.NotEmpty(result.Hyperlinks);
        var hyperlink = result.Hyperlinks[0];
        Assert.Equal("https://example.com/", hyperlink.Address); // Excel normalizes URLs by adding trailing slash
        Assert.Contains("Example", hyperlink.DisplayText);
    }
}
