using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Tests for PowerQuery refresh operations, including refresh with loadDestination parameter
/// </summary>
public partial class PowerQueryCommandsTests
{
    /// <summary>
    /// Verifies that refresh without loadDestination parameter maintains existing behavior
    /// (refreshes existing connections without changing configuration)
    /// </summary>
    [Fact]
    public async Task Refresh_WithoutLoadDestination_RefreshesExistingConnection()
    {
        // Arrange - Create a query loaded to worksheet
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests), 
            nameof(Refresh_WithoutLoadDestination_RefreshesExistingConnection), 
            _tempDir);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Refresh_WithoutLoadDestination_RefreshesExistingConnection));

        // Import query as worksheet (default loadDestination)
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestRefresh", testQueryFile, "worksheet");
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        await batch.SaveAsync();

        // Act - Refresh without loadDestination parameter
        await using var refreshBatch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var result = await _powerQueryCommands.RefreshAsync(refreshBatch, "TestRefresh");

        // Assert
        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.False(result.IsConnectionOnly, "Query should not be connection-only after refresh");
        Assert.Equal("TestRefresh", result.QueryName);
    }

    /// <summary>
    /// Verifies that refresh with loadDestination='worksheet' converts a connection-only query
    /// to a loaded query (applies load configuration then refreshes)
    /// </summary>
    [Fact]
    public async Task Refresh_WithLoadDestinationWorksheet_ConvertsConnectionOnlyToLoaded()
    {
        // Arrange - Create a connection-only query
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests), 
            nameof(Refresh_WithLoadDestinationWorksheet_ConvertsConnectionOnlyToLoaded), 
            _tempDir);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Refresh_WithLoadDestinationWorksheet_ConvertsConnectionOnlyToLoaded));

        // Import query as connection-only
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestQuery", testQueryFile, "connection-only");
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        await batch.SaveAsync();

        // Verify it's connection-only
        await using var verifyBatch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var configBefore = await _powerQueryCommands.GetLoadConfigAsync(verifyBatch, "TestQuery");
        Assert.True(configBefore.Success, "Failed to get config");
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, configBefore.LoadMode);

        // Act - Refresh with loadDestination='worksheet'
        // NOTE: This test validates the MCP Server layer behavior
        // For Core layer, we need to use SetLoadToTableAsync then RefreshAsync
        var setLoadResult = await _powerQueryCommands.SetLoadToTableAsync(verifyBatch, "TestQuery", "TestQuery");
        Assert.True(setLoadResult.Success, $"SetLoadToTable failed: {setLoadResult.ErrorMessage}");

        var refreshResult = await _powerQueryCommands.RefreshAsync(verifyBatch, "TestQuery");

        // Assert - Verify query is now loaded to worksheet
        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");
        Assert.False(refreshResult.IsConnectionOnly, "Query should no longer be connection-only");

        // Verify load configuration changed
        var configAfter = await _powerQueryCommands.GetLoadConfigAsync(verifyBatch, "TestQuery");
        Assert.True(configAfter.Success, "Failed to get config after refresh");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, configAfter.LoadMode);

        // Verify worksheet was created with data
        var sheetsResult = await _sheetCommands.ListAsync(verifyBatch);
        Assert.True(sheetsResult.Success);
        Assert.Contains(sheetsResult.Worksheets, w => w.Name == "TestQuery");

        // Verify QueryTable exists on worksheet
        await verifyBatch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item("TestQuery");
            dynamic queryTables = sheet.QueryTables;
            Assert.True(queryTables.Count > 0, "Expected at least one QueryTable on the worksheet");
            return 0;
        });
    }

    /// <summary>
    /// Verifies that refresh with loadDestination='data-model' converts a connection-only query
    /// to load into the Data Model
    /// </summary>
    [Fact]
    public async Task Refresh_WithLoadDestinationDataModel_LoadsToDataModel()
    {
        // Arrange - Create a connection-only query
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests), 
            nameof(Refresh_WithLoadDestinationDataModel_LoadsToDataModel), 
            _tempDir);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Refresh_WithLoadDestinationDataModel_LoadsToDataModel));

        // Import query as connection-only
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestDMQuery", testQueryFile, "connection-only");
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        await batch.SaveAsync();

        // Act - Apply load configuration and refresh
        await using var refreshBatch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var setLoadResult = await _powerQueryCommands.SetLoadToDataModelAsync(refreshBatch, "TestDMQuery");
        Assert.True(setLoadResult.Success, $"SetLoadToDataModel failed: {setLoadResult.ErrorMessage}");

        var refreshResult = await _powerQueryCommands.RefreshAsync(refreshBatch, "TestDMQuery");

        // Assert
        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");
        Assert.False(refreshResult.IsConnectionOnly, "Query should no longer be connection-only");

        // Verify query is in Data Model
        var dataModelCommands = new DataModelCommands();
        var tablesResult = await dataModelCommands.ListTablesAsync(refreshBatch);
        Assert.True(tablesResult.Success, $"Failed to list Data Model tables: {tablesResult.ErrorMessage}");
        Assert.Contains(tablesResult.Tables, t => t.Name == "TestDMQuery");
    }

    /// <summary>
    /// Verifies that refresh with loadDestination='both' loads query to both worksheet and Data Model
    /// </summary>
    [Fact]
    public async Task Refresh_WithLoadDestinationBoth_LoadsToBothDestinations()
    {
        // Arrange - Create a connection-only query
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests), 
            nameof(Refresh_WithLoadDestinationBoth_LoadsToBothDestinations), 
            _tempDir);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Refresh_WithLoadDestinationBoth_LoadsToBothDestinations));

        // Import query as connection-only
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestBothQuery", testQueryFile, "connection-only");
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        await batch.SaveAsync();

        // Act - Apply load configuration and refresh
        await using var refreshBatch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var setLoadResult = await _powerQueryCommands.SetLoadToBothAsync(refreshBatch, "TestBothQuery", "TestBothQuery");
        Assert.True(setLoadResult.Success, $"SetLoadToBoth failed: {setLoadResult.ErrorMessage}");

        var refreshResult = await _powerQueryCommands.RefreshAsync(refreshBatch, "TestBothQuery");

        // Assert
        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");
        Assert.False(refreshResult.IsConnectionOnly, "Query should no longer be connection-only");

        // Verify worksheet was created with data
        var sheetsResult = await _sheetCommands.ListAsync(refreshBatch);
        Assert.True(sheetsResult.Success);
        Assert.Contains(sheetsResult.Worksheets, w => w.Name == "TestBothQuery");

        // Verify QueryTable exists on worksheet
        await refreshBatch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item("TestBothQuery");
            dynamic queryTables = sheet.QueryTables;
            Assert.True(queryTables.Count > 0, "Expected at least one QueryTable on the worksheet");
            return 0;
        });

        // Verify query is in Data Model
        var dataModelCommands = new DataModelCommands();
        var tablesResult = await dataModelCommands.ListTablesAsync(refreshBatch);
        Assert.True(tablesResult.Success, $"Failed to list Data Model tables: {tablesResult.ErrorMessage}");
        Assert.Contains(tablesResult.Tables, t => t.Name == "TestBothQuery");
    }

    /// <summary>
    /// Verifies that refresh on a connection-only query without loadDestination parameter
    /// returns success with IsConnectionOnly=true (maintains connection-only state)
    /// </summary>
    [Fact]
    public async Task Refresh_ConnectionOnlyWithoutLoadDestination_RemainsConnectionOnly()
    {
        // Arrange - Create a connection-only query
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests), 
            nameof(Refresh_ConnectionOnlyWithoutLoadDestination_RemainsConnectionOnly), 
            _tempDir);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Refresh_ConnectionOnlyWithoutLoadDestination_RemainsConnectionOnly));

        // Import query as connection-only
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestQuery", testQueryFile, "connection-only");
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        await batch.SaveAsync();

        // Act - Refresh without loadDestination parameter
        await using var refreshBatch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var result = await _powerQueryCommands.RefreshAsync(refreshBatch, "TestQuery");

        // Assert - Query should remain connection-only
        Assert.True(result.Success, "Refresh should succeed for connection-only queries");
        Assert.True(result.IsConnectionOnly, "Query should remain connection-only when no loadDestination is specified");

        // Verify load configuration unchanged
        var configAfter = await _powerQueryCommands.GetLoadConfigAsync(refreshBatch, "TestQuery");
        Assert.True(configAfter.Success, "Failed to get config after refresh");
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, configAfter.LoadMode);
    }

    /// <summary>
    /// Verifies that specifying a custom targetSheet name works correctly with refresh
    /// </summary>
    [Fact]
    public async Task Refresh_WithCustomTargetSheet_CreatesSheetWithCorrectName()
    {
        // Arrange - Create a connection-only query
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests), 
            nameof(Refresh_WithCustomTargetSheet_CreatesSheetWithCorrectName), 
            _tempDir);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Refresh_WithCustomTargetSheet_CreatesSheetWithCorrectName));

        // Import query as connection-only
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestQuery", testQueryFile, "connection-only");
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        await batch.SaveAsync();

        // Act - Apply load configuration with custom sheet name
        await using var refreshBatch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var setLoadResult = await _powerQueryCommands.SetLoadToTableAsync(refreshBatch, "TestQuery", "CustomSheetName");
        Assert.True(setLoadResult.Success, $"SetLoadToTable failed: {setLoadResult.ErrorMessage}");

        var refreshResult = await _powerQueryCommands.RefreshAsync(refreshBatch, "TestQuery");

        // Assert
        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");

        // Verify custom sheet name was used
        var sheetsResult = await _sheetCommands.ListAsync(refreshBatch);
        Assert.True(sheetsResult.Success);
        Assert.Contains(sheetsResult.Worksheets, w => w.Name == "CustomSheetName");
        Assert.DoesNotContain(sheetsResult.Worksheets, w => w.Name == "TestQuery");
    }
}
