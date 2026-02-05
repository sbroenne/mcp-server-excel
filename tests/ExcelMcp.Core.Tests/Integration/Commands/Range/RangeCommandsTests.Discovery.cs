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

    [Fact]
    public void GetInfo_ValidRange_ReturnsGeometryInPoints()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Act - Get info for a range that has known geometry
        var result = _commands.GetInfo(batch, sheetName, "A1:C5");

        // Assert - Geometry should be populated (values vary by default column width/row height)
        Assert.True(result.Success);
        Assert.NotNull(result.Left);
        Assert.NotNull(result.Top);
        Assert.NotNull(result.Width);
        Assert.NotNull(result.Height);

        // All geometry values should be positive (in points)
        Assert.True(result.Left >= 0, "Left should be >= 0 points");
        Assert.True(result.Top >= 0, "Top should be >= 0 points");
        Assert.True(result.Width > 0, "Width should be > 0 points");
        Assert.True(result.Height > 0, "Height should be > 0 points");
    }

    [Fact]
    public void GetInfo_DifferentRanges_ReturnsDifferentGeometry()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Act - Get info for two different ranges
        var rangeA1 = _commands.GetInfo(batch, sheetName, "A1");
        var rangeB2 = _commands.GetInfo(batch, sheetName, "B2");

        // Assert - B2 should be offset from A1
        Assert.True(rangeA1.Success);
        Assert.True(rangeB2.Success);

        // B2 should have greater Left (offset by column A width)
        Assert.True(rangeB2.Left > rangeA1.Left, "B2 should be to the right of A1");

        // B2 should have greater Top (offset by row 1 height)
        Assert.True(rangeB2.Top > rangeA1.Top, "B2 should be below A1");
    }

    [Fact]
    public void GetInfo_LargerRange_ReturnsLargerDimensions()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Act - Compare single cell to multi-cell range
        var singleCell = _commands.GetInfo(batch, sheetName, "A1");
        var multiCell = _commands.GetInfo(batch, sheetName, "A1:C5");

        // Assert
        Assert.True(singleCell.Success);
        Assert.True(multiCell.Success);

        // Multi-cell range should be larger
        Assert.True(multiCell.Width > singleCell.Width, "A1:C5 should be wider than A1");
        Assert.True(multiCell.Height > singleCell.Height, "A1:C5 should be taller than A1");
    }

}




