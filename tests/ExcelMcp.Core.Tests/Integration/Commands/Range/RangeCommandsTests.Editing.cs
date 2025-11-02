using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for range editing operations
/// </summary>
public partial class RangeCommandsTests
{
    // === CLEAR OPERATIONS TESTS ===

    [Fact]
    public async Task ClearAllAsync_RemovesEverything()
    {
        // Arrange
        string testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), nameof(ClearAllAsync_RemovesEverything), _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        await _commands.SetValuesAsync(batch, "Sheet1", "A1", [new() { "Test" }]);

        // Act
        var result = await _commands.ClearAllAsync(batch, "Sheet1", "A1");
        // Assert
        Assert.True(result.Success);

        var readResult = await _commands.GetValuesAsync(batch, "Sheet1", "A1");
        Assert.Null(readResult.Values[0][0]);
    }

    [Fact]
    public async Task ClearContentsAsync_PreservesFormatting()
    {
        // Arrange
        string testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), nameof(ClearContentsAsync_PreservesFormatting), _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        await _commands.SetValuesAsync(batch, "Sheet1", "A1:B2",
        [
            new() { 1, 2 },
            new() { 3, 4 }
        ]);

        // Act
        var result = await _commands.ClearContentsAsync(batch, "Sheet1", "A1:B2");
        // Assert
        Assert.True(result.Success);

        var readResult = await _commands.GetValuesAsync(batch, "Sheet1", "A1:B2");
        Assert.All(readResult.Values, row => Assert.All(row, cell => Assert.Null(cell)));
    }

    // === COPY OPERATIONS TESTS ===

    [Fact]
    public async Task CopyAsync_CopiesRangeToNewLocation()
    {
        // Arrange
        string testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), nameof(CopyAsync_CopiesRangeToNewLocation), _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        var sourceData = new List<List<object?>>
        {
            new() { "A", "B" },
            new() { 1, 2 }
        };

        await _commands.SetValuesAsync(batch, "Sheet1", "A1:B2", sourceData);

        // Act
        var result = await _commands.CopyAsync(batch, "Sheet1", "A1:B2", "Sheet1", "D1:E2");
        // Assert
        Assert.True(result.Success);

        var readResult = await _commands.GetValuesAsync(batch, "Sheet1", "D1:E2");
        Assert.Equal("A", readResult.Values[0][0]);
        Assert.Equal(2.0, Convert.ToDouble(readResult.Values[1][1]));
    }

    [Fact]
    public async Task CopyValuesAsync_CopiesOnlyValues()
    {
        // Arrange
        string testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), nameof(CopyValuesAsync_CopiesOnlyValues), _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        await _commands.SetValuesAsync(batch, "Sheet1", "A1", [new() { 10 }]);
        await _commands.SetFormulasAsync(batch, "Sheet1", "B1", [new() { "=A1*2" }]);

        // Act
        var result = await _commands.CopyValuesAsync(batch, "Sheet1", "B1", "Sheet1", "C1");
        // Assert
        Assert.True(result.Success);

        // C1 should have value 20 but no formula
        var formulaResult = await _commands.GetFormulasAsync(batch, "Sheet1", "C1");
        Assert.Equal(20.0, Convert.ToDouble(formulaResult.Values[0][0]));
        Assert.Empty(formulaResult.Formulas[0][0]); // No formula
    }

    // === INSERT/DELETE OPERATIONS TESTS ===
}
