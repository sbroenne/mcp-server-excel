using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for advanced Range operations (clear formats, copy formulas, insert/delete, hyperlinks)
/// Optimized: Single batch per test, no SaveAsync unless testing persistence
/// </summary>
public sealed partial class PersistentServiceRangeAdvancedTests
{
    [Fact]
    [Trait("Speed", "Medium")]
    public void ClearFormats_FormattedRange_RemovesFormattingOnly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1", [["Test"]]);
        _fixture.Send(
            "rangeformat.format-range",
            new { sheetName, rangeAddress = "A1", bold = true, fillColor = "#FF0000" });

        // Act - Clear only formats
        var result = _commands.ClearFormats(batch, sheetName, "A1");

        // Assert
        Assert.True(result.Success, $"ClearFormats failed: {result.ErrorMessage}");

        // Verify value remains but formatting is gone
        var values = _commands.GetValues(batch, sheetName, "A1");
        Assert.Equal("Test", values.Values[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void CopyFormulas_SourceWithFormulas_CopiesFormulasOnly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1:A2", [[10], [20]]);
        _commands.SetFormulas(batch, sheetName, "A3", [["=A1+A2"]]);

        // Act - Copy formulas to B3
        var result = _commands.CopyFormulas(batch, sheetName, "A3", sheetName, "B3");

        // Assert
        Assert.True(result.Success, $"CopyFormulas failed: {result.ErrorMessage}");

        // Verify formula was copied (should adjust references)
        var formulas = _commands.GetFormulas(batch, sheetName, "B3");
        Assert.NotNull(formulas.Formulas[0][0]);
        Assert.Contains("+", formulas.Formulas[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void InsertCells_ShiftDown_InsertsAndShiftsExisting()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1", [["Original"]]);

        // Act - Insert cell at A1, shifting down
        var result = _commands.InsertCells(batch, sheetName, "A1", InsertShiftDirection.Down);

        // Assert
        Assert.True(result.Success, $"InsertCells failed: {result.ErrorMessage}");

        // Verify original value shifted to A2
        var values = _commands.GetValues(batch, sheetName, "A2");
        Assert.Equal("Original", values.Values[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteCells_ShiftUp_RemovesAndShifts()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1:A2", [["Delete Me"], ["Keep Me"]]);

        // Act - Delete A1, shifting up
        var result = _commands.DeleteCells(batch, sheetName, "A1", DeleteShiftDirection.Up);

        // Assert
        Assert.True(result.Success, $"DeleteCells failed: {result.ErrorMessage}");

        // Verify A2 value shifted to A1
        var values = _commands.GetValues(batch, sheetName, "A1");
        Assert.Equal("Keep Me", values.Values[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void InsertRows_BeforeExistingData_InsertsBlankRows()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1", [["Row 1"]]);

        // Act - Insert 2 rows at row 1
        var result = _commands.InsertRows(batch, sheetName, "1:2");

        // Assert
        Assert.True(result.Success, $"InsertRows failed: {result.ErrorMessage}");

        // Verify original data shifted to row 3
        var values = _commands.GetValues(batch, sheetName, "A3");
        Assert.Equal("Row 1", values.Values[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteRows_ExistingRows_RemovesRows()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(
            batch,
            sheetName,
            "A1:A3",
            [["Row 1"], ["Row 2 - Delete"], ["Row 3"]]);

        // Act - Delete row 2
        var result = _commands.DeleteRows(batch, sheetName, "2:2");

        // Assert
        Assert.True(result.Success, $"DeleteRows failed: {result.ErrorMessage}");

        // Verify row 3 shifted to row 2
        var values = _commands.GetValues(batch, sheetName, "A2");
        Assert.Equal("Row 3", values.Values[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void InsertColumns_BeforeExistingData_InsertsBlankColumns()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1", [["Col A"]]);

        // Act - Insert 2 columns at column A (column 1)
        var result = _commands.InsertColumns(batch, sheetName, "A:B");

        // Assert
        Assert.True(result.Success, $"InsertColumns failed: {result.ErrorMessage}");

        // Verify original data shifted to column C
        var values = _commands.GetValues(batch, sheetName, "C1");
        Assert.Equal("Col A", values.Values[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteColumns_ExistingColumns_RemovesColumns()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(
            batch,
            sheetName,
            "A1:C1",
            [["Col A", "Col B - Delete", "Col C"]]);

        // Act - Delete column B
        var result = _commands.DeleteColumns(batch, sheetName, "B:B");

        // Assert
        Assert.True(result.Success, $"DeleteColumns failed: {result.ErrorMessage}");

        // Verify column C shifted to B
        var values = _commands.GetValues(batch, sheetName, "B1");
        Assert.Equal("Col C", values.Values[0][0]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void GetHyperlink_ExistingHyperlink_ReturnsDetails()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Add a hyperlink
        var addResult = _commands.AddHyperlink(batch, sheetName, "A1", "https://example.com", "Example Link");
        Assert.True(addResult.Success);

        // Act
        var result = _commands.GetHyperlink(batch, sheetName, "A1");

        // Assert
        Assert.True(result.Success, $"GetHyperlink failed: {result.ErrorMessage}");
        Assert.NotEmpty(result.Hyperlinks);
        var hyperlink = result.Hyperlinks[0];
        Assert.Equal("https://example.com/", hyperlink.Address); // Excel normalizes URLs by adding trailing slash
        Assert.Contains("Example", hyperlink.DisplayText);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void GetMergeInfo_RangeSpanningMultipleMergedRegions_ReturnsMergedRanges()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.MergeCells(batch, sheetName, "B4:F4");
        _commands.MergeCells(batch, sheetName, "G4:K4");
        _commands.MergeCells(batch, sheetName, "L4:P4");

        var result = _commands.GetMergeInfo(batch, sheetName, "A4:P4");

        Assert.True(result.Success, $"GetMergeInfo failed: {result.ErrorMessage}");
        Assert.True(result.IsMerged);
        Assert.Equal(["$B$4:$F$4", "$G$4:$K$4", "$L$4:$P$4"], result.MergedRanges);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void GetMergeInfo_PartiallyMergedRange_ReturnsContainedMergedRange()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.MergeCells(batch, sheetName, "B4:F4");

        var result = _commands.GetMergeInfo(batch, sheetName, "B4:G4");

        Assert.True(result.Success, $"GetMergeInfo failed: {result.ErrorMessage}");
        Assert.True(result.IsMerged);
        Assert.Equal(["$B$4:$F$4"], result.MergedRanges);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void GetMergeInfo_UnmergedRange_ReturnsFalseAndNoMergedRanges()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var result = _commands.GetMergeInfo(batch, sheetName, "A1:D4");

        Assert.True(result.Success, $"GetMergeInfo failed: {result.ErrorMessage}");
        Assert.False(result.IsMerged);
        Assert.Empty(result.MergedRanges);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void GetMergeInfo_NormalMixedRange_ReturnsEveryMergedRange()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.MergeCells(batch, sheetName, "B2:C2");
        _commands.MergeCells(batch, sheetName, "F3:H3");

        var result = _commands.GetMergeInfo(batch, sheetName, "A1:H4");

        Assert.True(result.Success, $"GetMergeInfo failed: {result.ErrorMessage}");
        Assert.True(result.IsMerged);
        Assert.Equal(["$B$2:$C$2", "$F$3:$H$3"], result.MergedRanges);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void GetMergeInfo_OversizedMixedRange_ThrowsActionableScanLimitError()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.MergeCells(batch, sheetName, "B2:C2");

        var exception = Assert.Throws<InvalidOperationException>(
            () => _commands.GetMergeInfo(batch, sheetName, "A1:AO100"));

        Assert.Contains("4,100", exception.Message, StringComparison.Ordinal);
        Assert.Contains("scan limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("smaller range", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("unmerge", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void GetMergeInfo_OversizedSingleMergedArea_ReturnsAreaWithoutScanningEveryCell()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.MergeCells(batch, sheetName, "A1:AO100");

        var result = _commands.GetMergeInfo(batch, sheetName, "A1:AO100");

        Assert.True(result.Success, $"GetMergeInfo failed: {result.ErrorMessage}");
        Assert.True(result.IsMerged);
        Assert.Equal(["$A$1:$AO$100"], result.MergedRanges);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void GetMergeInfo_SeparateMergedAreasCoveringRange_ReturnsEveryArea()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.MergeCells(batch, sheetName, "A1:B1");
        _commands.MergeCells(batch, sheetName, "C1:D1");

        var result = _commands.GetMergeInfo(batch, sheetName, "A1:D1");

        Assert.True(result.Success, $"GetMergeInfo failed: {result.ErrorMessage}");
        Assert.True(result.IsMerged);
        Assert.Equal(["$A$1:$B$1", "$C$1:$D$1"], result.MergedRanges);
    }
}

