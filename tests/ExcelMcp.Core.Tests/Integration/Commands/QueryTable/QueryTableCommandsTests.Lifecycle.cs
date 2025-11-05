using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.QueryTable;

/// <summary>
/// Tests for QueryTable lifecycle operations (Create, Delete)
/// </summary>
public partial class QueryTableCommandsTests
{
    [Fact]
    public async Task CreateFromQuery_ValidQuery_CreatesQueryTable()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(CreateFromQuery_ValidQuery_CreatesQueryTable), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create a Power Query
        var dataModelCommands = new Sbroenne.ExcelMcp.Core.Commands.DataModelCommands();
        var pqCommands = new Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"Name\", \"Value\"}, {{\"A\", 1}, {\"B\", 2}}) in Source";
        var importResult = await pqCommands.ImportAsync(batch, "DataQuery", mCode);
        Assert.True(importResult.Success);
        
        // Create worksheet
        var sheetCommands = new Sbroenne.ExcelMcp.Core.Commands.SheetCommands();
        var createSheetResult = await sheetCommands.CreateAsync(batch, "Data");
        Assert.True(createSheetResult.Success);

        // Act
        var result = await _commands.CreateFromQueryAsync(batch, "Data", "MyQueryTable", "DataQuery");

        // Assert
        Assert.True(result.Success, $"Create failed: {result.ErrorMessage}");
        
        // Verify QueryTable exists
        var listResult = await _commands.ListAsync(batch);
        Assert.True(listResult.Success);
        var qt = Assert.Single(listResult.QueryTables);
        Assert.Equal("MyQueryTable", qt.Name);
        Assert.Equal("Data", qt.WorksheetName);
    }

    [Fact]
    public async Task CreateFromQuery_NonExistentQuery_ReturnsFalse()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(CreateFromQuery_NonExistentQuery_ReturnsFalse), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create worksheet
        var sheetCommands = new Sbroenne.ExcelMcp.Core.Commands.SheetCommands();
        await sheetCommands.CreateAsync(batch, "Data");

        // Act
        var result = await _commands.CreateFromQueryAsync(batch, "Data", "MyQueryTable", "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task CreateFromQuery_WithOptions_AppliesOptions()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(CreateFromQuery_WithOptions_AppliesOptions), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create Power Query and worksheet
        var dataModelCommands = new Sbroenne.ExcelMcp.Core.Commands.DataModelCommands();
        var pqCommands = new Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"Col1\"}, {{\"Val1\"}}) in Source";
        await pqCommands.ImportAsync(batch, "TestQuery", mCode);
        
        var sheetCommands = new Sbroenne.ExcelMcp.Core.Commands.SheetCommands();
        await sheetCommands.CreateAsync(batch, "Sheet1");

        var options = new QueryTableCreateOptions
        {
            BackgroundQuery = true,
            RefreshOnFileOpen = true,
            RefreshImmediately = true
        };

        // Act
        var result = await _commands.CreateFromQueryAsync(batch, "Sheet1", "TestQT", "TestQuery", "A1", options);

        // Assert
        Assert.True(result.Success, $"Create failed: {result.ErrorMessage}");
        
        // Verify properties
        var getResult = await _commands.GetAsync(batch, "TestQT");
        Assert.True(getResult.Success);
        Assert.NotNull(getResult.QueryTable);
        Assert.True(getResult.QueryTable.BackgroundQuery);
        Assert.True(getResult.QueryTable.RefreshOnFileOpen);
    }

    [Fact]
    public async Task Delete_ExistingQueryTable_DeletesSuccessfully()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(Delete_ExistingQueryTable_DeletesSuccessfully), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create QueryTable
        var dataModelCommands = new Sbroenne.ExcelMcp.Core.Commands.DataModelCommands();
        var pqCommands = new Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"A\"}, {{1}}) in Source";
        await pqCommands.ImportAsync(batch, "Q1", mCode);
        
        var sheetCommands = new Sbroenne.ExcelMcp.Core.Commands.SheetCommands();
        await sheetCommands.CreateAsync(batch, "S1");
        
        await _commands.CreateFromQueryAsync(batch, "S1", "QT1", "Q1");

        // Act
        var result = await _commands.DeleteAsync(batch, "QT1");

        // Assert
        Assert.True(result.Success, $"Delete failed: {result.ErrorMessage}");
        
        // Verify deleted
        var listResult = await _commands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Empty(listResult.QueryTables);
    }

    [Fact]
    public async Task Delete_NonExistentQueryTable_ReturnsFalse()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(Delete_NonExistentQueryTable_ReturnsFalse), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _commands.DeleteAsync(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
}
