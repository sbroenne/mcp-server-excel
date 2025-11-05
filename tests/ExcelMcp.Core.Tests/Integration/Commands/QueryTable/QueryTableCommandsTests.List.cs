using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.QueryTable;

/// <summary>
/// Tests for QueryTable List operations
/// </summary>
public partial class QueryTableCommandsTests
{
    [Fact]
    public async Task List_EmptyWorkbook_ReturnsSuccessWithEmptyList()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(List_EmptyWorkbook_ReturnsSuccessWithEmptyList), _tempDir);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _commands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.QueryTables);
        Assert.Empty(result.QueryTables);
    }

    [Fact]
    public async Task List_WithQueryTable_ReturnsQueryTable()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(List_WithQueryTable_ReturnsQueryTable), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // First create a simple Power Query - need to write M code to file
        var mCodeFile = Path.Combine(_tempDir, "TestQuery.pq");
        var mCode = "let Source = #table({\"Column1\"}, {{\"Value1\"}, {\"Value2\"}}) in Source";
        await System.IO.File.WriteAllTextAsync(mCodeFile, mCode);

        var dataModelCommands = new Sbroenne.ExcelMcp.Core.Commands.DataModelCommands();
        var pqCommands = new Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommands(dataModelCommands);
        var importResult = await pqCommands.ImportAsync(batch, "TestQuery", mCodeFile, loadDestination: "connection-only");
        Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");

        // Create a worksheet for the QueryTable
        var sheetCommands = new Sbroenne.ExcelMcp.Core.Commands.SheetCommands();
        var createSheetResult = await sheetCommands.CreateAsync(batch, "QuerySheet");
        Assert.True(createSheetResult.Success, $"Create sheet failed: {createSheetResult.ErrorMessage}");

        // Create QueryTable from the Power Query
        var createResult = await _commands.CreateFromQueryAsync(batch, "QuerySheet", "TestQueryTable", "TestQuery");
        Assert.True(createResult.Success, $"Create QueryTable failed: {createResult.ErrorMessage}");

        // Act - List QueryTables in same batch
        var result = await _commands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotEmpty(result.QueryTables);
        var qt = Assert.Single(result.QueryTables);
        Assert.Equal("TestQueryTable", qt.Name);
        Assert.Equal("QuerySheet", qt.WorksheetName);
    }
}
