using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for range discovery operations
/// </summary>
public partial class RangeCommandsTests
{
    // === NATIVE EXCEL COM OPERATIONS TESTS ===

    [Fact]
    public async Task GetUsedRange_SheetWithSparseData_ReturnsNonEmptyCells()
    {
        // Arrange
        string testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), nameof(GetUsedRange_SheetWithSparseData_ReturnsNonEmptyCells), _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        await _commands.SetValuesAsync(batch, "Sheet1", "A1", [new() { "Start" }]);
        await _commands.SetValuesAsync(batch, "Sheet1", "D10", [new() { "End" }]);

        // Act
        var result = await _commands.GetUsedRangeAsync(batch, "Sheet1");

        // Assert
        Assert.True(result.Success);
        Assert.True(result.RowCount >= 10);
        Assert.True(result.ColumnCount >= 4);
        Assert.Equal("Start", result.Values[0][0]);
    }

    [Fact]
    public async Task GetCurrentRegion_CellInPopulated3x3Range_ReturnsContiguousBlock()
    {
        // Arrange
        string testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), nameof(GetCurrentRegion_CellInPopulated3x3Range_ReturnsContiguousBlock), _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        await _commands.SetValuesAsync(batch, "Sheet1", "A1:C3",
        [
            new() { 1, 2, 3 },
            new() { 4, 5, 6 },
            new() { 7, 8, 9 }
        ]);

        // Act - Get region from middle cell
        var result = await _commands.GetCurrentRegionAsync(batch, "Sheet1", "B2");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(3, result.RowCount);
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal(1.0, Convert.ToDouble(result.Values[0][0]));
        Assert.Equal(9.0, Convert.ToDouble(result.Values[2][2]));
    }

    [Fact]
    public async Task GetRangeInfo_ValidAddress_ReturnsMetadata()
    {
        // Arrange
        string testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), nameof(GetRangeInfo_ValidAddress_ReturnsMetadata), _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        await _commands.SetValuesAsync(batch, "Sheet1", "A1:D10",
        [
            new() { 1, 2, 3, 4 }
        ]);

        // Act
        var result = await _commands.GetRangeInfoAsync(batch, "Sheet1", "A1:D10");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(10, result.RowCount);
        Assert.Equal(4, result.ColumnCount);
        Assert.Contains("$A$1:$D$10", result.Address); // Absolute address
    }

}
