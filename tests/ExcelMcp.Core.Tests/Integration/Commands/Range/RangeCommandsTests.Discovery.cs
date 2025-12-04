using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for range discovery operations
/// </summary>
public partial class RangeCommandsTests
{
    // === NATIVE EXCEL COM OPERATIONS TESTS ===

    [Fact]
    public void GetUsedRange_SheetWithSparseData_ReturnsNonEmptyCells()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1", [["Start"]]);
        _commands.SetValues(batch, sheetName, "D10", [["End"]]);

        // Act
        var result = _commands.GetUsedRange(batch, sheetName);

        // Assert
        Assert.True(result.Success);
        Assert.True(result.RowCount >= 10);
        Assert.True(result.ColumnCount >= 4);
        Assert.Equal("Start", result.Values[0][0]);
    }

    [Fact]
    public void GetCurrentRegion_CellInPopulated3x3Range_ReturnsContiguousBlock()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1:C3",
        [
            [1, 2, 3],
            [4, 5, 6],
            [7, 8, 9]
        ]);

        // Act - Get region from middle cell
        var result = _commands.GetCurrentRegion(batch, sheetName, "B2");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(3, result.RowCount);
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal(
            1.0,
            Convert.ToDouble(result.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            9.0,
            Convert.ToDouble(result.Values[2][2], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void GetInfo_ValidAddress_ReturnsMetadata()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1:D10",
        [
            [1, 2, 3, 4]
        ]);

        // Act
        var result = _commands.GetInfo(batch, sheetName, "A1:D10");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(10, result.RowCount);
        Assert.Equal(4, result.ColumnCount);
        Assert.Contains("$A$1:$D$10", result.Address); // Absolute address
    }

}
