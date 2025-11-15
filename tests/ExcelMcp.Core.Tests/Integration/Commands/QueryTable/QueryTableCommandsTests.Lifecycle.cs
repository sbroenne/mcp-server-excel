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
    /// <inheritdoc/>
    [Fact]
    public async Task CreateFromQuery_ValidQuery_CreatesQueryTable()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(CreateFromQuery_ValidQuery_CreatesQueryTable), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create a Power Query
        var dataModelCommands = new Core.Commands.DataModelCommands();
        var pqCommands = new Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"Name\", \"Value\"}, {{\"A\", 1}, {\"B\", 2}}) in Source";
        var mCodeFile = Path.Combine(_tempDir, "DataQuery.pq");
        await System.IO.File.WriteAllText(mCodeFile, mCode);
        var importResult = pqCommands.Create(batch, "DataQuery", mCodeFile, PowerQueryLoadMode.ConnectionOnly);
        Assert.True(importResult.Success);

        // Create worksheet
        var sheetCommands = new Core.Commands.SheetCommands();
        var createSheetResult = sheetCommands.Create(batch, "Data");
        Assert.True(createSheetResult.Success);

        // Act
        var result = _commands.CreateFromQuery(batch, "Data", "MyQueryTable", "DataQuery");

        // Assert
        Assert.True(result.Success, $"Create failed: {result.ErrorMessage}");

        // Verify QueryTable exists
        var listResult = _commands.List(batch);
        Assert.True(listResult.Success);
        var qt = Assert.Single(listResult.QueryTables);
        Assert.Equal("MyQueryTable", qt.Name);
        Assert.Equal("Data", qt.WorksheetName);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task CreateFromQuery_NonExistentQuery_ReturnsFalse()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(CreateFromQuery_NonExistentQuery_ReturnsFalse), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create worksheet
        var sheetCommands = new Core.Commands.SheetCommands();
        await sheetCommands.Create(batch, "Data");

        // Act
        var result = _commands.CreateFromQuery(batch, "Data", "MyQueryTable", "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task CreateFromQuery_WithOptions_AppliesOptions()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(CreateFromQuery_WithOptions_AppliesOptions), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create Power Query and worksheet
        var dataModelCommands = new Core.Commands.DataModelCommands();
        var pqCommands = new Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"Col1\"}, {{\"Val1\"}}) in Source";
        var mCodeFile = Path.Combine(_tempDir, "TestQuery.pq");
        await System.IO.File.WriteAllText(mCodeFile, mCode);
        await pqCommands.Create(batch, "TestQuery", mCodeFile, PowerQueryLoadMode.ConnectionOnly);

        var sheetCommands = new Core.Commands.SheetCommands();
        await sheetCommands.Create(batch, "Sheet1");

        var options = new QueryTableCreateOptions
        {
            BackgroundQuery = true,
            RefreshOnFileOpen = true,
            RefreshImmediately = true
        };

        // Act
        var result = _commands.CreateFromQuery(batch, "Sheet1", "TestQT", "TestQuery", "A1", options);

        // Assert
        Assert.True(result.Success, $"Create failed: {result.ErrorMessage}");

        // Verify properties
        var getResult = _commands.Get(batch, "TestQT");
        Assert.True(getResult.Success);
        Assert.NotNull(getResult.QueryTable);
        Assert.True(getResult.QueryTable.BackgroundQuery);
        Assert.True(getResult.QueryTable.RefreshOnFileOpen);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Delete_ExistingQueryTable_DeletesSuccessfully()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(Delete_ExistingQueryTable_DeletesSuccessfully), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create QueryTable
        var dataModelCommands = new Core.Commands.DataModelCommands();
        var pqCommands = new Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"A\"}, {{1}}) in Source";
        var mCodeFile = Path.Combine(_tempDir, "Q1.pq");
        await System.IO.File.WriteAllText(mCodeFile, mCode);
        await pqCommands.Create(batch, "Q1", mCodeFile, PowerQueryLoadMode.ConnectionOnly);

        var sheetCommands = new Core.Commands.SheetCommands();
        await sheetCommands.Create(batch, "S1");

        await _commands.CreateFromQuery(batch, "S1", "QT1", "Q1");

        // Act
        var result = _commands.Delete(batch, "QT1");

        // Assert
        Assert.True(result.Success, $"Delete failed: {result.ErrorMessage}");

        // Verify deleted
        var listResult = _commands.List(batch);
        Assert.True(listResult.Success);
        Assert.Empty(listResult.QueryTables);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Delete_NonExistentQueryTable_ReturnsFalse()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(Delete_NonExistentQueryTable_ReturnsFalse), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _commands.Delete(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
}




