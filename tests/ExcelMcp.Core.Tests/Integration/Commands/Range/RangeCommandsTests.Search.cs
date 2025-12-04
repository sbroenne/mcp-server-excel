using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for range search operations
/// </summary>
public partial class RangeCommandsTests
{
    // === FIND/REPLACE OPERATIONS TESTS ===

    [Fact]
    public void Find_FindsMatchingCells()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1:C2",
        [
            ["Apple", "Banana", "Apple"],
            ["Cherry", "Apple", "Banana"]
        ]);

        // Act
        var result = _commands.Find(batch, sheetName, "A1:C2", "Apple", new FindOptions
        {
            MatchCase = false,
            MatchEntireCell = true
        });

        // Assert
        Assert.True(result.Success);
        Assert.Equal(3, result.MatchingCells.Count); // Should find 3 "Apple" cells
    }

    [Fact]
    public void Replace_ReplacesAllOccurrences()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1:A3",
        [
            ["cat"],
            ["dog"],
            ["cat"]
        ]);

        // Act
        _commands.Replace(batch, sheetName, "A1:A3", "cat", "bird", new ReplaceOptions
        {
            ReplaceAll = true
        });

        // Assert - void method throws on failure, succeeds silently
        var readResult = _commands.GetValues(batch, sheetName, "A1:A3");
        Assert.Equal("bird", readResult.Values[0][0]);
        Assert.Equal("dog", readResult.Values[1][0]);
        Assert.Equal("bird", readResult.Values[2][0]);
    }

    // === SORT OPERATIONS TESTS ===

    [Fact]
    public void Sort_SortsRangeByColumn()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1:B4",
        [
            ["Name", "Age"],
            ["Charlie", 30],
            ["Alice", 25],
            ["Bob", 35]
        ]);

        // Act - Sort by first column (Name) ascending
        _commands.Sort(batch, sheetName, "A1:B4",
        [
            new() { ColumnIndex = 1, Ascending = true }
        ], hasHeaders: true);

        // Assert - void method throws on failure, succeeds silently
        var readResult = _commands.GetValues(batch, sheetName, "A2:A4");
        Assert.Equal("Alice", readResult.Values[0][0]);
        Assert.Equal("Bob", readResult.Values[1][0]);
        Assert.Equal("Charlie", readResult.Values[2][0]);
    }
}
