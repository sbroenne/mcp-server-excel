using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for range discovery operations
/// </summary>
public sealed partial class PersistentServiceRangeDiscoveryTests
{
    // === NATIVE EXCEL COM OPERATIONS TESTS ===

    [Fact]
    public async Task GetUsedRange_MissingSheet_ReturnsCategorizedNotFound()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, "B2", [["Keep"]]));
        var response = await _fixture.SendForFailureAsync(
            "range.get-used-range",
            new { sheetName = "MissingSheet" });

        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal("NotFound", response.ErrorCategory);
        var retained = _commands.GetValues(_fixture.BatchToken, sheet, "B2");
        RequireSuccess(retained);
        Assert.Equal("Keep", Assert.Single(Assert.Single(retained.Values)));
        var recovered = _commands.GetUsedRange(_fixture.BatchToken, sheet);
        RequireSuccess(recovered);
        Assert.Equal("$B$2", recovered.RangeAddress);
        Assert.Equal("Keep", Assert.Single(Assert.Single(recovered.Values)));
    }

    [Fact]
    public void GetUsedRange_SheetWithSparseData_ReturnsNonEmptyCells()
    {
        // Arrange - use shared file, create unique sheet for this test
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        Assert.True(_commands.SetValues(batch, sheetName, "A1", [["Start"]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "D10", [["End"]]).Success);

        // Act
        var result = _commands.GetUsedRange(batch, sheetName);

        // Assert
        Assert.True(result.Success);
        Assert.Equal(sheetName, result.SheetName);
        Assert.Equal("$A$1:$D$10", result.RangeAddress);
        Assert.Equal(10, result.RowCount);
        Assert.Equal(4, result.ColumnCount);
        Assert.Equal(10, result.Values.Count);
        for (var row = 0; row < 10; row++)
        {
            Assert.Equal(4, result.Values[row].Count);
            for (var column = 0; column < 4; column++)
            {
                object? expected = (row, column) switch
                {
                    (0, 0) => "Start",
                    (9, 3) => "End",
                    _ => null
                };
                Assert.Equal(expected, result.Values[row][column]);
            }
        }
    }

    [Fact]
    public void GetCurrentRegion_CellInPopulated3x3Range_ReturnsContiguousBlock()
    {
        // Arrange - use shared file, create unique sheet for this test
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        Assert.True(_commands.SetValues(batch, sheetName, "D5:F7",
        [
            [1, 2, 3],
            [4, 5, 6],
            [7, 8, 9]
        ]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [["Other region"]]).Success);

        // Act - Get region from middle cell
        var result = _commands.GetCurrentRegion(batch, sheetName, "E6");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(3, result.RowCount);
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal("$D$5:$F$7", result.RangeAddress);
        Assert.Equal(sheetName, result.SheetName);
        Assert.Equal(3, result.Values.Count);
        for (var row = 0; row < 3; row++)
        {
            Assert.Equal(3, result.Values[row].Count);
            for (var column = 0; column < 3; column++)
            {
                Assert.Equal(row * 3 + column + 1,
                    Convert.ToDouble(result.Values[row][column], System.Globalization.CultureInfo.InvariantCulture));
            }
        }
    }

    [Fact]
    public void GetInfo_ValidAddress_ReturnsMetadata()
    {
        // Arrange - use shared file, create unique sheet for this test
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        Assert.True(_commands.SetValues(batch, sheetName, "A1:D1",
        [
            [1, 2, 3, 4]
        ]).Success);

        // Act
        var result = _commands.GetInfo(batch, sheetName, "A1:D10");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(10, result.RowCount);
        Assert.Equal(4, result.ColumnCount);
        Assert.Equal("$A$1:$D$10", result.Address);
        Assert.Equal(sheetName, result.SheetName);
    }

    [Fact]
    public void GetInfo_ValidRange_ReturnsGeometryInPoints()
    {
        // Arrange - use shared file, create unique sheet for this test
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Act - Get info for a range that has known geometry
        var expected = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheetName, "B2:D6");
        var result = _commands.GetInfo(batch, sheetName, "B2:D6");

        // Assert - Geometry should be populated (values vary by default column width/row height)
        Assert.True(result.Success);
        Assert.Equal("$B$2:$D$6", result.Address);
        Assert.Equal(expected.Left, result.Left);
        Assert.Equal(expected.Top, result.Top);
        Assert.Equal(expected.Width, result.Width);
        Assert.Equal(expected.Height, result.Height);
    }

    [Fact]
    public void GetInfo_DifferentRanges_ReturnsDifferentGeometry()
    {
        // Arrange - use shared file, create unique sheet for this test
        var batch = _fixture.BatchToken;
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
        var expectedA1 = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheetName, "A1");
        var expectedB2 = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheetName, "B2");
        Assert.Equal(expectedA1.Left, rangeA1.Left);
        Assert.Equal(expectedA1.Top, rangeA1.Top);
        Assert.Equal(expectedB2.Left, rangeB2.Left);
        Assert.Equal(expectedB2.Top, rangeB2.Top);
        Assert.Equal(expectedA1.Width, rangeA1.Width);
        Assert.Equal(expectedA1.Height, rangeA1.Height);
        Assert.Equal(expectedB2.Width, rangeB2.Width);
        Assert.Equal(expectedB2.Height, rangeB2.Height);
    }

    [Fact]
    public void GetInfo_LargerRange_ReturnsLargerDimensions()
    {
        // Arrange - use shared file, create unique sheet for this test
        var batch = _fixture.BatchToken;
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
        var expectedSingle = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheetName, "A1");
        var expectedMulti = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheetName, "A1:C5");
        Assert.Equal(expectedSingle.Width, singleCell.Width);
        Assert.Equal(expectedSingle.Height, singleCell.Height);
        Assert.Equal(expectedMulti.Width, multiCell.Width);
        Assert.Equal(expectedMulti.Height, multiCell.Height);
        Assert.Equal(expectedMulti.Left, multiCell.Left);
        Assert.Equal(expectedMulti.Top, multiCell.Top);
    }

}
