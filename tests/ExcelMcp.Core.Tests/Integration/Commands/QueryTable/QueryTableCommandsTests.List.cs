using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.QueryTable;

/// <summary>
/// Tests for QueryTable List operations
/// </summary>
public partial class QueryTableCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public void List_EmptyWorkbook_ReturnsSuccessWithEmptyList()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(List_EmptyWorkbook_ReturnsSuccessWithEmptyList), _tempDir);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.List(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.QueryTables);
        Assert.Empty(result.QueryTables);
    }
    /// <inheritdoc/>

    [Fact]
    public void List_WithQueryTable_ReturnsQueryTable()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(List_WithQueryTable_ReturnsQueryTable), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // First create a simple Power Query
        var mCode = "let Source = #table({\"Column1\"}, {{\"Value1\"}, {\"Value2\"}}) in Source";

        var dataModelCommands = new DataModelCommands();
        var pqCommands = new PowerQueryCommands(dataModelCommands);
        var importResult = pqCommands.Create(batch, "TestQuery", mCode, PowerQueryLoadMode.ConnectionOnly);
        Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");

        // Create a worksheet for the QueryTable
        var sheetCommands = new SheetCommands();
        var createSheetResult = sheetCommands.Create(batch, "QuerySheet");
        Assert.True(createSheetResult.Success, $"Create sheet failed: {createSheetResult.ErrorMessage}");

        // Create QueryTable from the Power Query
        var createResult = _commands.CreateFromQuery(batch, "QuerySheet", "TestQueryTable", "TestQuery");
        Assert.True(createResult.Success, $"Create QueryTable failed: {createResult.ErrorMessage}");

        // Act - List QueryTables in same batch
        var result = _commands.List(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotEmpty(result.QueryTables);
        var qt = Assert.Single(result.QueryTables);
        Assert.Equal("TestQueryTable", qt.Name);
        Assert.Equal("QuerySheet", qt.WorksheetName);
    }
}



