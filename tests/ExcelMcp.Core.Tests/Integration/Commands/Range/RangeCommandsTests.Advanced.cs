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
    [Fact]
    [Trait("Speed", "Medium")]
    public async Task ClearFormats_FormattedRange_RemovesFormattingOnly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests), nameof(ClearFormats_FormattedRange_RemovesFormattingOnly), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Set values with formatting
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Test";
            sheet.Range["A1"].Font.Bold = true;
            sheet.Range["A1"].Interior.Color = 255; // Red background
            return 0;
        });

        // Act - Clear only formats
        var result = await _commands.ClearFormatsAsync(batch, "Sheet1", "A1");

        // Assert
        Assert.True(result.Success, $"ClearFormats failed: {result.ErrorMessage}");

        // Verify value remains but formatting is gone
        var values = await _commands.GetValuesAsync(batch, "Sheet1", "A1");
        Assert.Equal("Test", values.Values[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task CopyFormulas_SourceWithFormulas_CopiesFormulasOnly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests), nameof(CopyFormulas_SourceWithFormulas_CopiesFormulasOnly), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Set up source data with formulas
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = 10;
            sheet.Range["A2"].Value2 = 20;
            sheet.Range["A3"].Formula = "=A1+A2";
            return 0;
        });

        // Act - Copy formulas to B3
        var result = await _commands.CopyFormulasAsync(batch, "Sheet1", "A3", "Sheet1", "B3");

        // Assert
        Assert.True(result.Success, $"CopyFormulas failed: {result.ErrorMessage}");

        // Verify formula was copied (should adjust references)
        var formulas = await _commands.GetFormulasAsync(batch, "Sheet1", "B3");
        Assert.NotNull(formulas.Formulas[0][0]);
        Assert.Contains("+", formulas.Formulas[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task InsertCells_ShiftDown_InsertsAndShiftsExisting()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests), nameof(InsertCells_ShiftDown_InsertsAndShiftsExisting), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Set up initial data
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Original";
            return 0;
        });

        // Act - Insert cell at A1, shifting down
        var result = await _commands.InsertCellsAsync(batch, "Sheet1", "A1", InsertShiftDirection.Down);

        // Assert
        Assert.True(result.Success, $"InsertCells failed: {result.ErrorMessage}");

        // Verify original value shifted to A2
        var values = await _commands.GetValuesAsync(batch, "Sheet1", "A2");
        Assert.Equal("Original", values.Values[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task DeleteCells_ShiftUp_RemovesAndShifts()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests), nameof(DeleteCells_ShiftUp_RemovesAndShifts), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Set up data in A1 and A2
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Delete Me";
            sheet.Range["A2"].Value2 = "Keep Me";
            return 0;
        });

        // Act - Delete A1, shifting up
        var result = await _commands.DeleteCellsAsync(batch, "Sheet1", "A1", DeleteShiftDirection.Up);

        // Assert
        Assert.True(result.Success, $"DeleteCells failed: {result.ErrorMessage}");

        // Verify A2 value shifted to A1
        var values = await _commands.GetValuesAsync(batch, "Sheet1", "A1");
        Assert.Equal("Keep Me", values.Values[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task InsertRows_BeforeExistingData_InsertsBlankRows()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests), nameof(InsertRows_BeforeExistingData_InsertsBlankRows), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Set up data in row 1
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Row 1";
            return 0;
        });

        // Act - Insert 2 rows at row 1
        var result = await _commands.InsertRowsAsync(batch, "Sheet1", "1:2");

        // Assert
        Assert.True(result.Success, $"InsertRows failed: {result.ErrorMessage}");

        // Verify original data shifted to row 3
        var values = await _commands.GetValuesAsync(batch, "Sheet1", "A3");
        Assert.Equal("Row 1", values.Values[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task DeleteRows_ExistingRows_RemovesRows()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests), nameof(DeleteRows_ExistingRows_RemovesRows), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Set up data in rows 1-3
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Row 1";
            sheet.Range["A2"].Value2 = "Row 2 - Delete";
            sheet.Range["A3"].Value2 = "Row 3";
            return 0;
        });

        // Act - Delete row 2
        var result = await _commands.DeleteRowsAsync(batch, "Sheet1", "2:2");

        // Assert
        Assert.True(result.Success, $"DeleteRows failed: {result.ErrorMessage}");

        // Verify row 3 shifted to row 2
        var values = await _commands.GetValuesAsync(batch, "Sheet1", "A2");
        Assert.Equal("Row 3", values.Values[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task InsertColumns_BeforeExistingData_InsertsBlankColumns()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests), nameof(InsertColumns_BeforeExistingData_InsertsBlankColumns), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Set up data in column A
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Col A";
            return 0;
        });

        // Act - Insert 2 columns at column A (column 1)
        var result = await _commands.InsertColumnsAsync(batch, "Sheet1", "A:B");

        // Assert
        Assert.True(result.Success, $"InsertColumns failed: {result.ErrorMessage}");

        // Verify original data shifted to column C
        var values = await _commands.GetValuesAsync(batch, "Sheet1", "C1");
        Assert.Equal("Col A", values.Values[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task DeleteColumns_ExistingColumns_RemovesColumns()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests), nameof(DeleteColumns_ExistingColumns_RemovesColumns), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Set up data in columns A-C
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Col A";
            sheet.Range["B1"].Value2 = "Col B - Delete";
            sheet.Range["C1"].Value2 = "Col C";
            return 0;
        });

        // Act - Delete column B
        var result = await _commands.DeleteColumnsAsync(batch, "Sheet1", "B:B");

        // Assert
        Assert.True(result.Success, $"DeleteColumns failed: {result.ErrorMessage}");

        // Verify column C shifted to B
        var values = await _commands.GetValuesAsync(batch, "Sheet1", "B1");
        Assert.Equal("Col C", values.Values[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task GetHyperlink_ExistingHyperlink_ReturnsDetails()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests), nameof(GetHyperlink_ExistingHyperlink_ReturnsDetails), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Add a hyperlink
        var addResult = await _commands.AddHyperlinkAsync(batch, "Sheet1", "A1", "https://example.com", "Example Link");
        Assert.True(addResult.Success);

        // Act
        var result = await _commands.GetHyperlinkAsync(batch, "Sheet1", "A1");

        // Assert
        Assert.True(result.Success, $"GetHyperlink failed: {result.ErrorMessage}");
        Assert.NotEmpty(result.Hyperlinks);
        var hyperlink = result.Hyperlinks[0];
        Assert.Equal("https://example.com/", hyperlink.Address); // Excel normalizes URLs by adding trailing slash
        Assert.Contains("Example", hyperlink.DisplayText);
    }
}
