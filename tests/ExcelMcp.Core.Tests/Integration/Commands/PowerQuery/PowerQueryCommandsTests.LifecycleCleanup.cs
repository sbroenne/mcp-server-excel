using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Tests for Power Query lifecycle cleanup operations.
///
/// These tests validate that:
/// - Unload correctly removes Data Model connections
/// - Delete correctly removes Data Model connections (no orphans)
/// - List correctly reports IsConnectionOnly for Data Model queries
/// </summary>
/// <remarks>
/// Created for GitHub Issue #279: Power Query Lifecycle bugs
/// </remarks>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("Feature", "DataModel")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public class PowerQueryLifecycleCleanupTests : IClassFixture<TempDirectoryFixture>
{
    private readonly PowerQueryCommands _powerQueryCommands;
    private readonly DataModelCommands _dataModelCommands;
    private readonly ConnectionCommands _connectionCommands;
    private readonly TempDirectoryFixture _fixture;

    public PowerQueryLifecycleCleanupTests(TempDirectoryFixture fixture)
    {
        _dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(_dataModelCommands);
        _connectionCommands = new ConnectionCommands();
        _fixture = fixture;
    }

    #region Unload Tests - Data Model Connection Cleanup

    /// <summary>
    /// Issue #279 Fix 1: Unload with LoadToDataModel should remove the Data Model connection.
    /// Before fix: Connection remained orphaned after unload.
    /// </summary>
    [Fact]
    public async Task Unload_DataModelOnly_RemovesDataModelConnection()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_UnloadDM_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query with LoadToDataModel
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToDataModel);

        // Verify Data Model table exists
        var tablesBefore = await _dataModelCommands.ListTables(batch);
        Assert.True(tablesBefore.Success);
        Assert.Contains(tablesBefore.Tables, t => t.Name == queryName);

        // Verify connection exists (pattern: "Query - {queryName}")
        var connsBefore = _connectionCommands.List(batch);
        Assert.True(connsBefore.Success);
        Assert.Contains(connsBefore.Connections, c => c.Name.Contains($"Query - {queryName}"));

        // Act - Unload the query
        var unloadResult = _powerQueryCommands.Unload(batch, queryName);

        // Assert
        Assert.True(unloadResult.Success, $"Unload failed: {unloadResult.ErrorMessage}");

        // Verify Data Model connection is removed
        var connsAfter = _connectionCommands.List(batch);
        Assert.True(connsAfter.Success);
        Assert.DoesNotContain(connsAfter.Connections, c => c.Name.Contains($"Query - {queryName}"));

        // Verify query still exists (just unloaded, not deleted)
        var queries = _powerQueryCommands.List(batch);
        Assert.True(queries.Success);
        Assert.Contains(queries.Queries, q => q.Name == queryName);

        // Verify query is now ConnectionOnly
        var loadConfig = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfig.Success);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, loadConfig.LoadMode);
    }

    /// <summary>
    /// Issue #279 Fix 1: Unload with LoadToBoth should remove both worksheet data AND Data Model connection.
    /// </summary>
    [Fact]
    public async Task Unload_LoadToBoth_RemovesBothWorksheetAndDataModelConnection()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_UnloadBoth_" + Guid.NewGuid().ToString("N")[..8];
        var sheetName = "BothSheet";
        var mCode = @"let Source = #table({""Val""}, {{42}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query with LoadToBoth
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToBoth, sheetName);

        // Verify Data Model table exists
        var tablesBefore = await _dataModelCommands.ListTables(batch);
        Assert.True(tablesBefore.Success);
        Assert.Contains(tablesBefore.Tables, t => t.Name == queryName);

        // Verify connection exists
        var connsBefore = _connectionCommands.List(batch);
        Assert.True(connsBefore.Success);
        Assert.Contains(connsBefore.Connections, c => c.Name.Contains($"Query - {queryName}"));

        // Act - Unload the query
        var unloadResult = _powerQueryCommands.Unload(batch, queryName);

        // Assert
        Assert.True(unloadResult.Success, $"Unload failed: {unloadResult.ErrorMessage}");

        // Verify Data Model connection is removed
        var connsAfter = _connectionCommands.List(batch);
        Assert.True(connsAfter.Success);
        Assert.DoesNotContain(connsAfter.Connections, c => c.Name.Contains($"Query - {queryName}"));

        // Verify query is now ConnectionOnly
        var loadConfig = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfig.Success);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, loadConfig.LoadMode);
    }

    #endregion

    #region Delete Tests - Data Model Connection Cleanup

    /// <summary>
    /// Issue #279 Fix 3: Delete with LoadToDataModel should remove the Data Model connection.
    /// Before fix: Connection remained orphaned after delete.
    /// </summary>
    [Fact]
    public async Task Delete_DataModelOnly_RemovesDataModelConnection()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_DeleteDM_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query with LoadToDataModel
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToDataModel);

        // Verify Data Model table exists
        var tablesBefore = await _dataModelCommands.ListTables(batch);
        Assert.True(tablesBefore.Success);
        Assert.Contains(tablesBefore.Tables, t => t.Name == queryName);

        // Verify connection exists
        var connsBefore = _connectionCommands.List(batch);
        Assert.True(connsBefore.Success);
        Assert.Contains(connsBefore.Connections, c => c.Name.Contains($"Query - {queryName}"));

        // Act - Delete the query
        _powerQueryCommands.Delete(batch, queryName);

        // Assert - Verify query is gone
        var queries = _powerQueryCommands.List(batch);
        Assert.True(queries.Success);
        Assert.DoesNotContain(queries.Queries, q => q.Name == queryName);

        // Verify Data Model connection is removed (no orphans)
        var connsAfter = _connectionCommands.List(batch);
        Assert.True(connsAfter.Success);
        Assert.DoesNotContain(connsAfter.Connections, c => c.Name.Contains($"Query - {queryName}"));
    }

    /// <summary>
    /// Issue #279 Fix 3: Delete with LoadToBoth should remove both worksheet data AND Data Model connection.
    /// </summary>
    [Fact]
    public async Task Delete_LoadToBoth_RemovesBothWorksheetAndDataModelConnection()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_DeleteBoth_" + Guid.NewGuid().ToString("N")[..8];
        var sheetName = "DeleteBothSheet";
        var mCode = @"let Source = #table({""Val""}, {{99}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query with LoadToBoth
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToBoth, sheetName);

        // Verify Data Model table exists
        var tablesBefore = await _dataModelCommands.ListTables(batch);
        Assert.True(tablesBefore.Success);
        Assert.Contains(tablesBefore.Tables, t => t.Name == queryName);

        // Verify connection exists
        var connsBefore = _connectionCommands.List(batch);
        Assert.True(connsBefore.Success);
        Assert.Contains(connsBefore.Connections, c => c.Name.Contains($"Query - {queryName}"));

        // Act - Delete the query
        _powerQueryCommands.Delete(batch, queryName);

        // Assert - Verify query is gone
        var queries = _powerQueryCommands.List(batch);
        Assert.True(queries.Success);
        Assert.DoesNotContain(queries.Queries, q => q.Name == queryName);

        // Verify Data Model connection is removed (no orphans)
        var connsAfter = _connectionCommands.List(batch);
        Assert.True(connsAfter.Success);
        Assert.DoesNotContain(connsAfter.Connections, c => c.Name.Contains($"Query - {queryName}"));
    }

    #endregion

    #region List IsConnectionOnly Tests - Data Model Awareness

    /// <summary>
    /// Issue #279 Fix 2: List should NOT report IsConnectionOnly=true for Data Model queries.
    /// Before fix: Queries loaded ONLY to Data Model were reported as ConnectionOnly.
    /// </summary>
    [Fact]
    public void List_DataModelOnly_NotReportedAsConnectionOnly()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_ListDM_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query with LoadToDataModel
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToDataModel);

        // Act - List queries
        var listResult = _powerQueryCommands.List(batch);

        // Assert
        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
        var query = listResult.Queries.FirstOrDefault(q => q.Name == queryName);
        Assert.NotNull(query);

        // CRITICAL: LoadToDataModel should NOT be reported as ConnectionOnly
        Assert.False(query.IsConnectionOnly,
            "Query loaded to Data Model should NOT be reported as ConnectionOnly");
    }

    /// <summary>
    /// List should correctly report ConnectionOnly for queries with no load destination.
    /// </summary>
    [Fact]
    public void List_ConnectionOnly_ReportsAsConnectionOnly()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_ListConnOnly_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query with ConnectionOnly
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Act - List queries
        var listResult = _powerQueryCommands.List(batch);

        // Assert
        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
        var query = listResult.Queries.FirstOrDefault(q => q.Name == queryName);
        Assert.NotNull(query);
        Assert.True(query.IsConnectionOnly,
            "Query with ConnectionOnly should be reported as ConnectionOnly");
    }

    /// <summary>
    /// List should NOT report IsConnectionOnly=true for queries loaded to worksheet.
    /// </summary>
    [Fact]
    public void List_LoadToTable_NotReportedAsConnectionOnly()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_ListTable_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query with LoadToTable
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToTable);

        // Act - List queries
        var listResult = _powerQueryCommands.List(batch);

        // Assert
        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
        var query = listResult.Queries.FirstOrDefault(q => q.Name == queryName);
        Assert.NotNull(query);
        Assert.False(query.IsConnectionOnly,
            "Query loaded to table should NOT be reported as ConnectionOnly");
    }

    /// <summary>
    /// List should NOT report IsConnectionOnly=true for LoadToBoth queries.
    /// </summary>
    [Fact]
    public void List_LoadToBoth_NotReportedAsConnectionOnly()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_ListBoth_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query with LoadToBoth
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToBoth);

        // Act - List queries
        var listResult = _powerQueryCommands.List(batch);

        // Assert
        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
        var query = listResult.Queries.FirstOrDefault(q => q.Name == queryName);
        Assert.NotNull(query);
        Assert.False(query.IsConnectionOnly,
            "Query loaded to both should NOT be reported as ConnectionOnly");
    }

    /// <summary>
    /// Mixed scenario: Correctly identify ConnectionOnly among multiple query types.
    /// </summary>
    [Fact]
    public void List_MixedLoadModes_CorrectlyIdentifiesConnectionOnlyQueries()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var suffix = Guid.NewGuid().ToString("N")[..6];
        var queryConnOnly = "PQ_Mix_ConnOnly_" + suffix;
        var queryTable = "PQ_Mix_Table_" + suffix;
        var queryDataModel = "PQ_Mix_DataModel_" + suffix;
        var queryBoth = "PQ_Mix_Both_" + suffix;

        var mCode = @"let Source = #table({""A""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create queries with different load modes
        _powerQueryCommands.Create(batch, queryConnOnly, mCode, PowerQueryLoadMode.ConnectionOnly);
        _powerQueryCommands.Create(batch, queryTable, mCode, PowerQueryLoadMode.LoadToTable, "Sheet1");
        _powerQueryCommands.Create(batch, queryDataModel, mCode, PowerQueryLoadMode.LoadToDataModel);
        _powerQueryCommands.Create(batch, queryBoth, mCode, PowerQueryLoadMode.LoadToBoth, "Sheet2");

        // Act
        var listResult = _powerQueryCommands.List(batch);

        // Assert
        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");

        var connOnlyQuery = listResult.Queries.FirstOrDefault(q => q.Name == queryConnOnly);
        var tableQuery = listResult.Queries.FirstOrDefault(q => q.Name == queryTable);
        var dataModelQuery = listResult.Queries.FirstOrDefault(q => q.Name == queryDataModel);
        var bothQuery = listResult.Queries.FirstOrDefault(q => q.Name == queryBoth);

        Assert.NotNull(connOnlyQuery);
        Assert.NotNull(tableQuery);
        Assert.NotNull(dataModelQuery);
        Assert.NotNull(bothQuery);

        // ONLY ConnectionOnly should report IsConnectionOnly=true
        Assert.True(connOnlyQuery.IsConnectionOnly, "ConnectionOnly query should be ConnectionOnly");
        Assert.False(tableQuery.IsConnectionOnly, "LoadToTable should NOT be ConnectionOnly");
        Assert.False(dataModelQuery.IsConnectionOnly, "LoadToDataModel should NOT be ConnectionOnly");
        Assert.False(bothQuery.IsConnectionOnly, "LoadToBoth should NOT be ConnectionOnly");
    }

    #endregion
}
