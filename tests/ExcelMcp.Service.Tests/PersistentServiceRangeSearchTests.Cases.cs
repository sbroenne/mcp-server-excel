using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for range search operations
/// </summary>
public sealed partial class PersistentServiceRangeSearchTests
{
    // === FIND/REPLACE OPERATIONS TESTS ===

    [Fact]
    public void Find_DefaultLimit_ReturnsOnlyTenOfTwentyFiveMatches()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        List<List<object?>> values = Enumerable.Range(0, 26)
            .Select(index => new List<object?> { index < 25 ? "Apple" : "Banana" })
            .ToList();
        var setResult = _commands.SetValues(batch, sheetName, "A1:A26", values);
        Assert.True(setResult.Success, setResult.ErrorMessage);

        var result = _commands.Find(batch, sheetName, "A1:A26", "Apple", new FindOptions
        {
            MatchEntireCell = true
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(10, result.MatchingCells.Count);
        AssertFindCoverage(result, 25, 10);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(10)]
    [InlineData(11)]
    [InlineData(25)]
    public void Find_OmittedLimit_ReturnsExactCoverage(int matchCount)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var rangeAddress = $"A1:A{matchCount + 1}";
        List<List<object?>> values = Enumerable.Range(0, matchCount + 1)
            .Select(index => new List<object?> { index < matchCount ? "Apple" : "Banana" })
            .ToList();
        Assert.True(_commands.SetValues(batch, sheetName, rangeAddress, values).Success);

        var response = _fixture.Send("rangeedit.find", new
        {
            sheetName,
            rangeAddress,
            searchValue = "Apple",
            findOptions = new FindOptions { MatchEntireCell = true }
        });

        using var document = JsonDocument.Parse(response.Result!);
        var root = document.RootElement;
        var returnedCount = Math.Min(matchCount, 10);
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.Equal(matchCount, root.GetProperty("totalCount").GetInt64());
        Assert.Equal(returnedCount, root.GetProperty("returnedCount").GetInt32());
        Assert.Equal(matchCount > returnedCount, root.GetProperty("truncated").GetBoolean());
        Assert.Equal(returnedCount, root.GetProperty("matchingCells").GetArrayLength());
        var result = ServiceCommandProxy.DeserializeResult<RangeFindResult>(response.Result!);
        AssertFindCoverage(result, matchCount, returnedCount);
        AssertReturnedAppleCells(result, matchCount);
    }

    [Theory]
    [InlineData(1, 1)]
    [InlineData(5, 5)]
    [InlineData(25, 25)]
    [InlineData(26, 25)]
    [InlineData(1001, 25)]
    [InlineData(int.MaxValue, 25)]
    public void Find_ExplicitLimit_ReturnsRequestedCoverage(int maxMatches, int returnedCount)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        List<List<object?>> values = Enumerable.Range(0, 26)
            .Select(index => new List<object?> { index < 25 ? "Apple" : "Banana" })
            .ToList();
        Assert.True(_commands.SetValues(batch, sheetName, "A1:A26", values).Success);

        var result = _commands.Find(
            batch, sheetName, "A1:A26", "Apple",
            new FindOptions { MatchEntireCell = true }, maxMatches);

        AssertFindCoverage(result, 25, returnedCount);
        AssertReturnedAppleCells(result, 25);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(int.MinValue)]
    public void Find_NonpositiveLimit_ReportsInvalidInput(int maxMatches)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            _commands.Find(_fixture.BatchToken, sheetName, "A1", "Apple", new FindOptions(), maxMatches));

        Assert.Contains("maxMatches", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("1.5")]
    [InlineData("2147483648")]
    public async Task Find_NonIntegerOrOverflowLimit_ReportsInvalidInput(string maxMatchesJson)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var response = await _fixture.SendForFailureAsync("rangeedit.find", new
        {
            sheetName,
            rangeAddress = "A1",
            searchValue = "Apple",
            findOptions = new FindOptions(),
            maxMatches = JsonSerializer.Deserialize<JsonElement>(maxMatchesJson)
        });

        Assert.False(response.Success);
        Assert.Contains("maxMatches", response.ErrorMessage, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, false, 3)]
    [InlineData(true, false, 2)]
    [InlineData(false, true, 2)]
    [InlineData(true, true, 1)]
    public void Find_LimitedSearch_PreservesCaseAndWholeCellMatching(
        bool matchCase, bool matchEntireCell, int totalCount)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:A4",
            [["Apple"], ["apple"], ["Apple pie"], ["Banana"]]).Success);

        var result = _commands.Find(batch, sheetName, "A1:A4", "Apple", new FindOptions
        {
            MatchCase = matchCase,
            MatchEntireCell = matchEntireCell
        }, maxMatches: 1);

        AssertFindCoverage(result, totalCount, 1);
        Assert.All(result.MatchingCells, cell => Assert.InRange(cell.Row, 1, 3));
    }

    [Theory]
    [InlineData(true, false, "=1+1", 1)]
    [InlineData(false, true, "2", 2)]
    [InlineData(true, true, "2", 2)]
    public void Find_LimitedSearch_PreservesFormulaAndValueSelection(
        bool searchFormulas, bool searchValues, string searchValue, int totalCount)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetFormulas(batch, sheetName, "A1:A2",
            [["=1+1"], ["=2"]]).Success);

        var result = _commands.Find(batch, sheetName, "A1:A2", searchValue, new FindOptions
        {
            SearchFormulas = searchFormulas,
            SearchValues = searchValues,
            MatchEntireCell = true
        }, maxMatches: 1);

        AssertFindCoverage(result, totalCount, 1);
        Assert.Equal(2, Assert.Single(result.MatchingCells).Value);
    }

    private static void AssertFindCoverage(RangeFindResult result, long totalCount, int returnedCount)
    {
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal(totalCount, result.TotalCount);
        Assert.Equal(returnedCount, result.ReturnedCount);
        Assert.Equal(returnedCount, result.MatchingCells.Count);
        Assert.Equal(totalCount > returnedCount, result.Truncated);
        Assert.Equal(returnedCount, result.MatchingCells.Select(cell => cell.Address).Distinct().Count());
    }

    private static void AssertReturnedAppleCells(RangeFindResult result, int matchCount)
    {
        Assert.All(result.MatchingCells, cell =>
        {
            Assert.InRange(cell.Row, 1, matchCount);
            Assert.Equal(1, cell.Column);
            Assert.Equal($"$A${cell.Row}", cell.Address);
            Assert.Equal("Apple", cell.Value);
        });
    }

    [Fact]
    public void Find_FindsMatchingCells()
    {
        // Arrange
        var batch = _fixture.BatchToken;
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
        AssertFindCoverage(result, 3, 3);
    }

    [Fact]
    public void Replace_ReplacesAllOccurrences()
    {
        // Arrange
        var batch = _fixture.BatchToken;
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
        var batch = _fixture.BatchToken;
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


