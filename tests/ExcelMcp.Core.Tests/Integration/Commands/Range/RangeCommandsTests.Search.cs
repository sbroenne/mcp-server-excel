using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for range search operations
/// </summary>
public partial class RangeCommandsTests
{
    // === FIND/REPLACE OPERATIONS TESTS ===

    [Fact]
    public async Task Find_FindsMatchingCells()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), $"{Guid.NewGuid():N}", _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        await _commands.SetValuesAsync(batch, "Sheet1", "A1:C2",
        [
            new() { "Apple", "Banana", "Apple" },
            new() { "Cherry", "Apple", "Banana" }
        ]);

        // Act
        var result = await _commands.FindAsync(batch, "Sheet1", "A1:C2", "Apple", new FindOptions
        {
            MatchCase = false,
            MatchEntireCell = true
        });

        // Assert
        Assert.True(result.Success);
        Assert.Equal(3, result.MatchingCells.Count); // Should find 3 "Apple" cells
    }

    [Fact]
    public async Task Replace_ReplacesAllOccurrences()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), $"{Guid.NewGuid():N}", _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        await _commands.SetValuesAsync(batch, "Sheet1", "A1:A3",
        [
            new() { "cat" },
            new() { "dog" },
            new() { "cat" }
        ]);

        // Act
        var result = await _commands.ReplaceAsync(batch, "Sheet1", "A1:A3", "cat", "bird", new ReplaceOptions
        {
            ReplaceAll = true
        });
        // Assert
        Assert.True(result.Success);

        var readResult = await _commands.GetValuesAsync(batch, "Sheet1", "A1:A3");
        Assert.Equal("bird", readResult.Values[0][0]);
        Assert.Equal("dog", readResult.Values[1][0]);
        Assert.Equal("bird", readResult.Values[2][0]);
    }

    // === SORT OPERATIONS TESTS ===

    [Fact]
    public async Task Sort_SortsRangeByColumn()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), $"{Guid.NewGuid():N}", _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        await _commands.SetValuesAsync(batch, "Sheet1", "A1:B4",
        [
            new() { "Name", "Age" },
            new() { "Charlie", 30 },
            new() { "Alice", 25 },
            new() { "Bob", 35 }
        ]);

        // Act - Sort by first column (Name) ascending
        var result = await _commands.SortAsync(batch, "Sheet1", "A1:B4",
        [
            new() { ColumnIndex = 1, Ascending = true }
        ], hasHeaders: true);
        // Assert
        if (!result.Success)
        {
            _output.WriteLine($"Sort failed: {result.ErrorMessage}");
        }
        Assert.True(result.Success);

        var readResult = await _commands.GetValuesAsync(batch, "Sheet1", "A2:A4");
        Assert.Equal("Alice", readResult.Values[0][0]);
        Assert.Equal("Bob", readResult.Values[1][0]);
        Assert.Equal("Charlie", readResult.Values[2][0]);
    }

}
