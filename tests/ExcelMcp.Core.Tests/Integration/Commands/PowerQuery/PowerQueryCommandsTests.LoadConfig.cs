using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Tests for PowerQuery load configuration operations
/// </summary>
public partial class PowerQueryCommandsTests
{
    [Fact]
    public async Task SetConnectionOnly_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(PowerQueryCommandsTests), nameof(SetConnectionOnly_WithExistingQuery_ReturnsSuccessResult), _tempDir);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(SetConnectionOnly_WithExistingQuery_ReturnsSuccessResult));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestConnectionOnly", testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        var result = await _powerQueryCommands.SetConnectionOnlyAsync(batch, "TestConnectionOnly");
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"SetConnectionOnly failed: {result.ErrorMessage}");
        Assert.Equal("pq-set-connection-only", result.Action);
    }

    [Fact]
    public async Task SetLoadToTable_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(PowerQueryCommandsTests), nameof(SetLoadToTable_WithExistingQuery_ReturnsSuccessResult), _tempDir);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(SetLoadToTable_WithExistingQuery_ReturnsSuccessResult));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestLoadToTable", testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        var result = await _powerQueryCommands.SetLoadToTableAsync(batch, "TestLoadToTable", "TestSheet");

        // Assert - Verify the operation succeeded
        Assert.True(result.Success, $"SetLoadToTable failed: {result.ErrorMessage}");
        Assert.Equal("pq-set-load-to-table", result.Action);

        // Verify the load configuration was actually set
        var configResult = await _powerQueryCommands.GetLoadConfigAsync(batch, "TestLoadToTable");
        Assert.True(configResult.Success, $"Failed to get load config: {configResult.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, configResult.LoadMode);

        // Verify sheet was created
        var sheetsResult = await _sheetCommands.ListAsync(batch);
        Assert.True(sheetsResult.Success);
        Assert.Contains(sheetsResult.Worksheets, w => w.Name == "TestSheet");

        // Verify table/QueryTable exists on worksheet (actual data loaded)
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item("TestSheet");
            dynamic queryTables = sheet.QueryTables;
            Assert.True(queryTables.Count > 0, "Expected at least one QueryTable on the worksheet");
            return 0;
        });

        // Verify query is NOT in Data Model (LoadToTable only)
        var dataModelCommands = new DataModelCommands();
        var tablesResult = await dataModelCommands.ListTablesAsync(batch);
        Assert.True(tablesResult.Success, $"Failed to list Data Model tables: {tablesResult.ErrorMessage}");
        Assert.DoesNotContain(tablesResult.Tables, t => t.Name == "TestLoadToTable");

        await batch.SaveAsync();
    }

    [Fact]
    public async Task SetLoadToDataModel_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(PowerQueryCommandsTests), nameof(SetLoadToDataModel_WithExistingQuery_ReturnsSuccessResult), _tempDir);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(SetLoadToDataModel_WithExistingQuery_ReturnsSuccessResult));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestLoadToDataModel", testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        var result = await _powerQueryCommands.SetLoadToDataModelAsync(batch, "TestLoadToDataModel");

        // Assert - Verify the operation succeeded
        Assert.True(result.Success, $"SetLoadToDataModel failed: {result.ErrorMessage}");
        Assert.Equal("pq-set-load-to-data-model", result.Action);

        // Verify the load configuration was actually set
        var configResult = await _powerQueryCommands.GetLoadConfigAsync(batch, "TestLoadToDataModel");
        Assert.True(configResult.Success, $"Failed to get load config: {configResult.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, configResult.LoadMode);

        // Verify query is in Data Model
        var dataModelCommands = new DataModelCommands();
        var tablesResult = await dataModelCommands.ListTablesAsync(batch);
        Assert.True(tablesResult.Success, $"Failed to list Data Model tables: {tablesResult.ErrorMessage}");
        Assert.Contains(tablesResult.Tables, t => t.Name == "TestLoadToDataModel");

        // Verify NO QueryTable on any worksheet (LoadToDataModel only, no worksheet table)
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheets = ctx.Book.Worksheets;
            int sheetCount = sheets.Count;
            for (int i = 1; i <= sheetCount; i++)
            {
                dynamic sheet = sheets.Item(i);
                dynamic queryTables = sheet.QueryTables;
                Assert.True(queryTables.Count == 0, $"Expected no QueryTables on sheet '{sheet.Name}' for LoadToDataModel mode");
            }
            return 0;
        });

        await batch.SaveAsync();
    }

    [Fact]
    public async Task SetLoadToBoth_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(PowerQueryCommandsTests), nameof(SetLoadToBoth_WithExistingQuery_ReturnsSuccessResult), _tempDir);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(SetLoadToBoth_WithExistingQuery_ReturnsSuccessResult));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestLoadToBoth", testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        var result = await _powerQueryCommands.SetLoadToBothAsync(batch, "TestLoadToBoth", "TestSheet");

        // Assert - Verify the operation succeeded
        Assert.True(result.Success, $"SetLoadToBoth failed: {result.ErrorMessage}");
        Assert.Equal("pq-set-load-to-both", result.Action);

        // Verify sheet was created with a table
        var sheetsResult = await _sheetCommands.ListAsync(batch);
        Assert.True(sheetsResult.Success);
        Assert.Contains(sheetsResult.Worksheets, w => w.Name == "TestSheet");

        // Verify table exists on worksheet (QueryTable from Power Query)
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item("TestSheet");
            dynamic queryTables = sheet.QueryTables;
            Assert.True(queryTables.Count > 0, "Expected at least one QueryTable on the worksheet");
            return 0;
        });

        // Verify query is in Data Model
        var dataModelCommands = new DataModelCommands();
        var tablesResult = await dataModelCommands.ListTablesAsync(batch);
        Assert.True(tablesResult.Success, $"Failed to list Data Model tables: {tablesResult.ErrorMessage}");
        Assert.Contains(tablesResult.Tables, t => t.Name == "TestLoadToBoth");

        await batch.SaveAsync();
    }
}
