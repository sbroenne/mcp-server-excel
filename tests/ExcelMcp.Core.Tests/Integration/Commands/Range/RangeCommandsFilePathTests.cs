using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Range;

/// <summary>
/// Tests for filePath-based Range API (Phase 1 proof-of-concept)
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Ranges")]
[Trait("RequiresExcel", "true")]
public sealed class RangeCommandsFilePathTests : IClassFixture<TempDirectoryFixture>
{
    private readonly RangeCommands _commands;
    private readonly WorkbookCommands _workbookCommands;
    private readonly string _tempDir;

    public RangeCommandsFilePathTests(TempDirectoryFixture fixture)
    {
        _commands = new RangeCommands();
        _workbookCommands = new WorkbookCommands();
        _tempDir = fixture.TempDir;
    }

    [Fact]
    public async Task GetValues_FilePathBased_ReturnsValues()
    {
        // Arrange - Create test file with data
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsFilePathTests),
            nameof(GetValues_FilePathBased_ReturnsValues),
            _tempDir,
            ".xlsx");

        // Use batch API to set up initial data
        await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
        {
            await _commands.SetValuesAsync(batch, "Sheet1", "A1:B2",
            [
                ["Header1", "Header2"],
                ["Value1", "Value2"]
            ]);
            await batch.SaveAsync();
        }

        // Act - Use filePath-based API to read data
        var result = await _commands.GetValuesAsync(testFile, "Sheet1", "A1:B2");

        // Assert
        Assert.True(result.Success, $"GetValues failed: {result.ErrorMessage}");
        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal("Header1", result.Values[0][0]);
        Assert.Equal("Header2", result.Values[0][1]);
        Assert.Equal("Value1", result.Values[1][0]);
        Assert.Equal("Value2", result.Values[1][1]);

        // Cleanup
        await _workbookCommands.CloseAsync(testFile);
    }

    [Fact]
    public async Task GetValues_SequentialCallsReuseHandle_Succeeds()
    {
        // Arrange - Create test file
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsFilePathTests),
            nameof(GetValues_SequentialCallsReuseHandle_Succeeds),
            _tempDir,
            ".xlsx");

        // Set up initial data
        await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
        {
            await _commands.SetValuesAsync(batch, "Sheet1", "A1", [[42]]);
            await batch.SaveAsync();
        }

        // Act - Make multiple sequential calls using filePath-based API
        // This should reuse the cached handle automatically
        var result1 = await _commands.GetValuesAsync(testFile, "Sheet1", "A1");
        var result2 = await _commands.GetValuesAsync(testFile, "Sheet1", "A1");
        var result3 = await _commands.GetValuesAsync(testFile, "Sheet1", "A1");

        // Assert - All calls should succeed
        Assert.True(result1.Success, $"First call failed: {result1.ErrorMessage}");
        Assert.True(result2.Success, $"Second call failed: {result2.ErrorMessage}");
        Assert.True(result3.Success, $"Third call failed: {result3.ErrorMessage}");

        // All should return the same value
        Assert.Equal(42, Convert.ToInt32(result1.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(42, Convert.ToInt32(result2.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(42, Convert.ToInt32(result3.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));

        // Cleanup
        await _workbookCommands.CloseAsync(testFile);
    }

    [Fact]
    public async Task MultipleFiles_EachCachedSeparately_BothAccessible()
    {
        // Arrange - Create two test files
        var testFile1 = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsFilePathTests),
            $"{nameof(MultipleFiles_EachCachedSeparately_BothAccessible)}_File1",
            _tempDir,
            ".xlsx");

        var testFile2 = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsFilePathTests),
            $"{nameof(MultipleFiles_EachCachedSeparately_BothAccessible)}_File2",
            _tempDir,
            ".xlsx");

        // Set up data in both files
        await using (var batch1 = await ExcelSession.BeginBatchAsync(testFile1))
        {
            await _commands.SetValuesAsync(batch1, "Sheet1", "A1", [["File1Data"]]);
            await batch1.SaveAsync();
        }

        await using (var batch2 = await ExcelSession.BeginBatchAsync(testFile2))
        {
            await _commands.SetValuesAsync(batch2, "Sheet1", "A1", [["File2Data"]]);
            await batch2.SaveAsync();
        }

        // Act - Access both files using filePath-based API
        var result1 = await _commands.GetValuesAsync(testFile1, "Sheet1", "A1");
        var result2 = await _commands.GetValuesAsync(testFile2, "Sheet1", "A1");

        // Assert - Both files accessible, each returns correct data
        Assert.True(result1.Success, $"File1 access failed: {result1.ErrorMessage}");
        Assert.True(result2.Success, $"File2 access failed: {result2.ErrorMessage}");

        Assert.Equal("File1Data", result1.Values[0][0]?.ToString());
        Assert.Equal("File2Data", result2.Values[0][0]?.ToString());

        // Cleanup
        await _workbookCommands.CloseAsync(testFile1);
        await _workbookCommands.CloseAsync(testFile2);
    }
}
