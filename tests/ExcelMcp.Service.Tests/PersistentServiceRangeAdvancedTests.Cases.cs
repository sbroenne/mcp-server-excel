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

        RequireSuccess(_commands.SetValues(batch, sheetName, "A1", [["Test"]]));
        var normal = PersistentServiceRangeVerification.ReadFormat(_fixture, sheetName, "B1");
        _fixture.Send(
            "rangeformat.format",
            new
            {
                sheetName,
                rangeAddresses = (string[])["A1"],
                formatOptions = new
                {
                    bold = true,
                    fillColor = "#FF0000"
                }
            });
        Assert.NotEqual(normal, PersistentServiceRangeVerification.ReadFormat(_fixture, sheetName, "A1"));

        // Act - Clear only formats
        var result = _commands.ClearFormats(batch, sheetName, "A1");

        // Assert
        Assert.True(result.Success, $"ClearFormats failed: {result.ErrorMessage}");

        // Verify value remains but formatting is gone
        var values = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(values.Success, values.ErrorMessage);
        Assert.Equal("Test", values.Values[0][0]?.ToString());
        Assert.Equal(normal, PersistentServiceRangeVerification.ReadFormat(_fixture, sheetName, "A1"));
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void CopyFormulas_SourceWithFormulas_CopiesFormulasOnly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        RequireSuccess(_commands.SetValues(batch, sheetName, "A1:A2", [[10], [20]]));
        Assert.True(_commands.SetValues(batch, sheetName, "B1:B2", [[5], [6]]).Success);
        RequireSuccess(_commands.SetFormulas(batch, sheetName, "A3", [["=A1+A2"]]));
        Assert.True(_commands.SetNumberFormat(batch, sheetName, "B3", "0.00%").Success);
        var destinationFormat = PersistentServiceRangeVerification.ReadFormat(_fixture, sheetName, "B3");

        // Act - Copy formulas to B3
        var result = _commands.Copy(batch, sheetName, "A3", sheetName, "B3", PasteKind.Formulas);

        // Assert
        Assert.True(result.Success, $"CopyFormulas failed: {result.ErrorMessage}");

        // Verify formula was copied (should adjust references)
        var formulas = _commands.GetFormulas(batch, sheetName, "B3");
        Assert.True(formulas.Success, formulas.ErrorMessage);
        Assert.Equal("=B1+B2", formulas.Formulas[0][0]);
        Assert.Equal(11d, Convert.ToDouble(formulas.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(destinationFormat, PersistentServiceRangeVerification.ReadFormat(_fixture, sheetName, "B3"));
        var source = RequireSuccess(_commands.GetFormulas(batch, sheetName, "A3"));
        Assert.Equal("=A1+A2", Assert.Single(Assert.Single(source.Formulas)));
        Assert.Equal(30d, Convert.ToDouble(Assert.Single(Assert.Single(source.Values)),
            System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void InsertCells_ShiftDown_InsertsAndShiftsExisting()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        RequireSuccess(_commands.SetValues(batch, sheetName, "A1:B2",
            [["Original", "Neighbor 1"], ["Below", "Neighbor 2"]]));

        // Act - Insert cell at A1, shifting down
        var result = _commands.InsertCells(batch, sheetName, "A1", InsertShiftDirection.Down);

        // Assert
        Assert.True(result.Success, $"InsertCells failed: {result.ErrorMessage}");

        // Verify original value shifted to A2
        AssertCells(sheetName, "A1:B3",
            [[null, "Neighbor 1"], ["Original", "Neighbor 2"], ["Below", null]]);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteCells_ShiftUp_RemovesAndShifts()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        RequireSuccess(_commands.SetValues(batch, sheetName, "A1:B3",
            [["Delete Me", "Neighbor 1"], ["Keep Me", "Neighbor 2"], ["Below", "Neighbor 3"]]));

        // Act - Delete A1, shifting up
        var result = _commands.DeleteCells(batch, sheetName, "A1", DeleteShiftDirection.Up);

        // Assert
        Assert.True(result.Success, $"DeleteCells failed: {result.ErrorMessage}");

        // Verify A2 value shifted to A1
        AssertCells(sheetName, "A1:B3",
            [["Keep Me", "Neighbor 1"], ["Below", "Neighbor 2"], [null, "Neighbor 3"]]);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void InsertRows_BeforeExistingData_InsertsBlankRows()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        RequireSuccess(_commands.SetValues(batch, sheetName, "A1:B2",
            [["Row 1", "First"], ["Row 2", "Second"]]));

        // Act - Insert 2 rows at row 1
        var result = _commands.InsertRows(batch, sheetName, "1:2");

        // Assert
        Assert.True(result.Success, $"InsertRows failed: {result.ErrorMessage}");

        // Verify original data shifted to row 3
        AssertCells(sheetName, "A1:B4",
            [[null, null], [null, null], ["Row 1", "First"], ["Row 2", "Second"]]);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteRows_ExistingRows_RemovesRows()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        RequireSuccess(_commands.SetValues(
            batch,
            sheetName,
            "A1:B3",
            [["Row 1", "First"], ["Row 2 - Delete", "Removed"], ["Row 3", "Third"]]));

        // Act - Delete row 2
        var result = _commands.DeleteRows(batch, sheetName, "2:2");

        // Assert
        Assert.True(result.Success, $"DeleteRows failed: {result.ErrorMessage}");

        // Verify row 3 shifted to row 2
        AssertCells(sheetName, "A1:B3",
            [["Row 1", "First"], ["Row 3", "Third"], [null, null]]);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void InsertColumns_BeforeExistingData_InsertsBlankColumns()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        RequireSuccess(_commands.SetValues(batch, sheetName, "A1:B2",
            [["Col A", "Col B"], ["Second A", "Second B"]]));

        // Act - Insert 2 columns at column A (column 1)
        var result = _commands.InsertColumns(batch, sheetName, "A:B");

        // Assert
        Assert.True(result.Success, $"InsertColumns failed: {result.ErrorMessage}");

        // Verify original data shifted to column C
        AssertCells(sheetName, "A1:D2",
            [[null, null, "Col A", "Col B"], [null, null, "Second A", "Second B"]]);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteColumns_ExistingColumns_RemovesColumns()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        RequireSuccess(_commands.SetValues(
            batch,
            sheetName,
            "A1:C2",
            [["Col A", "Col B - Delete", "Col C"], ["Second A", "Removed", "Second C"]]));

        // Act - Delete column B
        var result = _commands.DeleteColumns(batch, sheetName, "B:B");

        // Assert
        Assert.True(result.Success, $"DeleteColumns failed: {result.ErrorMessage}");

        // Verify column C shifted to B
        AssertCells(sheetName, "A1:C2",
            [["Col A", "Col C", null], ["Second A", "Second C", null]]);
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
        var hyperlink = Assert.Single(result.Hyperlinks);
        Assert.Equal("https://example.com/", hyperlink.Address); // Excel normalizes URLs by adding trailing slash
        Assert.Equal("Example Link", hyperlink.DisplayText);
        Assert.Equal("A1", hyperlink.CellAddress);
        AssertCells(sheetName, "A1", [["Example Link"]]);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void GetMergeInfo_RangeSpanningMultipleMergedRegions_ReturnsMergedRanges()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        RequireSuccess(_commands.MergeCells(batch, sheetName, "B4:F4"));
        RequireSuccess(_commands.MergeCells(batch, sheetName, "G4:K4"));
        RequireSuccess(_commands.MergeCells(batch, sheetName, "L4:P4"));

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

        RequireSuccess(_commands.MergeCells(batch, sheetName, "B4:F4"));

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

        RequireSuccess(_commands.MergeCells(batch, sheetName, "B2:C2"));
        RequireSuccess(_commands.MergeCells(batch, sheetName, "F3:H3"));

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

        RequireSuccess(_commands.SetValues(batch, sheetName, "B2", [["Retained anchor"]]));
        RequireSuccess(_commands.MergeCells(batch, sheetName, "B2:C2"));

        var exception = Assert.Throws<InvalidOperationException>(
            () => _commands.GetMergeInfo(batch, sheetName, "A1:AO100"));

        Assert.Contains("4,100", exception.Message, StringComparison.Ordinal);
        Assert.Contains("scan limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("smaller range", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("unmerge", exception.Message, StringComparison.OrdinalIgnoreCase);
        var preserved = RequireSuccess(_commands.GetMergeInfo(batch, sheetName, "B2:C2"));
        Assert.Equal(["$B$2:$C$2"], preserved.MergedRanges);
        AssertCells(sheetName, "B2", [["Retained anchor"]]);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void GetMergeInfo_OversizedSingleMergedArea_ReturnsAreaWithoutScanningEveryCell()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        RequireSuccess(_commands.MergeCells(batch, sheetName, "A1:AO100"));

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

        RequireSuccess(_commands.MergeCells(batch, sheetName, "A1:B1"));
        RequireSuccess(_commands.MergeCells(batch, sheetName, "C1:D1"));

        var result = _commands.GetMergeInfo(batch, sheetName, "A1:D1");

        Assert.True(result.Success, $"GetMergeInfo failed: {result.ErrorMessage}");
        Assert.True(result.IsMerged);
        Assert.Equal(["$A$1:$B$1", "$C$1:$D$1"], result.MergedRanges);
    }

    [Theory]
    [InlineData("insert-cells")]
    [InlineData("delete-cells")]
    [InlineData("insert-rows")]
    [InlineData("delete-rows")]
    [InlineData("insert-columns")]
    [InlineData("delete-columns")]
    public void Editing_InvalidRange_PreservesValuesFormulasAndFormatting(string action)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheetName, "A1:B2",
            [["Retained", "Neighbor"], ["Below", "Unchanged"]]));
        RequireSuccess(_commands.SetFormulas(batch, sheetName, "C1", [["=6*7"]]));
        RequireSuccess(_commands.Format(batch, sheetName, ["A1:C2"],
            new() { Bold = true, FillColor = "#FFFF00" }));
        var before = PersistentServiceRangeVerification.ReadFormat(_fixture, sheetName, "A1");

        var error = Assert.Throws<InvalidOperationException>(() =>
        {
            const string address = "NotARange!!";
            _ = action switch
            {
                "insert-cells" => _commands.InsertCells(batch, sheetName, address, InsertShiftDirection.Down),
                "delete-cells" => _commands.DeleteCells(batch, sheetName, address, DeleteShiftDirection.Up),
                "insert-rows" => _commands.InsertRows(batch, sheetName, address),
                "delete-rows" => _commands.DeleteRows(batch, sheetName, address),
                "insert-columns" => _commands.InsertColumns(batch, sheetName, address),
                "delete-columns" => _commands.DeleteColumns(batch, sheetName, address),
                _ => throw new ArgumentOutOfRangeException(nameof(action))
            };
        });

        Assert.Contains("NotARange!!", error.Message);
        AssertCells(sheetName, "A1:B2", [["Retained", "Neighbor"], ["Below", "Unchanged"]]);
        var formula = RequireSuccess(_commands.GetFormulas(batch, sheetName, "C1"));
        Assert.Equal("=6*7", Assert.Single(Assert.Single(formula.Formulas)));
        Assert.Equal(42d, Convert.ToDouble(Assert.Single(Assert.Single(formula.Values)),
            System.Globalization.CultureInfo.InvariantCulture));
        foreach (var address in new[] { "A1", "B1", "C1", "A2", "B2", "C2" })
        {
            Assert.Equal(before, PersistentServiceRangeVerification.ReadFormat(_fixture, sheetName, address));
        }
    }

    private void AssertCells(string sheetName, string address, List<List<object?>> expected)
    {
        var result = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, address));
        Assert.Equal(expected.Count, result.Values.Count);
        for (var row = 0; row < expected.Count; row++)
        {
            Assert.Equal(expected[row], result.Values[row]);
        }
    }
}
