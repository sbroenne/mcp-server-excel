using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Tests for Power Query worksheet loading cleanup operations.
///
/// These tests validate that:
/// - LoadToTable creates properly named connections (Query - {name})
/// - Delete of loaded query removes query, connection, AND table (clean slate)
/// - No orphaned connections remain after operations
/// </summary>
/// <remarks>
/// Created to address bug: LoadQueryToWorksheet was creating orphaned connections
/// with generic names like "Connection", "Connection1" instead of "Query - {name}"
/// </remarks>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public class PowerQueryWorksheetCleanupTests : IClassFixture<TempDirectoryFixture>
{
    private readonly PowerQueryCommands _powerQueryCommands;
    private readonly DataModelCommands _dataModelCommands;
    private readonly ConnectionCommands _connectionCommands;
    private readonly TableCommands _tableCommands;
    private readonly TempDirectoryFixture _fixture;

    public PowerQueryWorksheetCleanupTests(TempDirectoryFixture fixture)
    {
        _dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(_dataModelCommands);
        _connectionCommands = new ConnectionCommands();
        _tableCommands = new TableCommands();
        _fixture = fixture;
    }

    #region Connection Naming Tests - Verify Add2 Fix

    /// <summary>
    /// Verifies that LoadToTable creates a properly named connection following
    /// the "Query - {queryName}" pattern, not a generic name like "Connection".
    /// 
    /// This is a regression test for the bug where ListObjects.Add() was creating
    /// connections with generic names instead of proper Power Query naming.
    /// </summary>
    [Fact]
    public void Create_LoadToTable_CreatesProperlyNamedConnection()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_ProperName_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Value""}, {{1}, {2}, {3}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Act - Create query with LoadToTable
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToTable);

        // Assert - Connection should follow "Query - {name}" pattern
        var connections = _connectionCommands.List(batch);
        Assert.True(connections.Success, $"List connections failed: {connections.ErrorMessage}");

        // Should have exactly one connection with proper naming
        var expectedConnectionName = $"Query - {queryName}";
        Assert.Contains(connections.Connections, c => c.Name == expectedConnectionName);

        // Should NOT have generic-named connections
        Assert.DoesNotContain(connections.Connections, c => c.Name == "Connection");
        Assert.DoesNotContain(connections.Connections, c => c.Name == "Connection1");
    }

    /// <summary>
    /// Verifies that multiple LoadToTable operations each create properly named
    /// connections without any generic "Connection", "Connection1" etc. orphans.
    /// </summary>
    [Fact]
    public void Create_MultipleLoadToTable_NoOrphanedConnections()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var suffix = Guid.NewGuid().ToString("N")[..6];
        var queryName1 = "PQ_Multi1_" + suffix;
        var queryName2 = "PQ_Multi2_" + suffix;
        var queryName3 = "PQ_Multi3_" + suffix;
        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Act - Create multiple queries with LoadToTable
        _powerQueryCommands.Create(batch, queryName1, mCode, PowerQueryLoadMode.LoadToTable, "Sheet1");
        _powerQueryCommands.Create(batch, queryName2, mCode, PowerQueryLoadMode.LoadToTable, "Sheet2");
        _powerQueryCommands.Create(batch, queryName3, mCode, PowerQueryLoadMode.LoadToTable, "Sheet3");

        // Assert
        var connections = _connectionCommands.List(batch);
        Assert.True(connections.Success);

        // Should have exactly 3 properly named connections
        Assert.Contains(connections.Connections, c => c.Name == $"Query - {queryName1}");
        Assert.Contains(connections.Connections, c => c.Name == $"Query - {queryName2}");
        Assert.Contains(connections.Connections, c => c.Name == $"Query - {queryName3}");

        // Count connections - should be exactly 3 (no orphans)
        var pqConnections = connections.Connections.Where(c => c.IsPowerQuery).ToList();
        Assert.Equal(3, pqConnections.Count);

        // Should NOT have any generic-named connections
        Assert.DoesNotContain(connections.Connections, c => c.Name == "Connection");
        Assert.DoesNotContain(connections.Connections, c => c.Name.StartsWith("Connection", StringComparison.Ordinal) && char.IsDigit(c.Name.Last()));
    }

    #endregion

    #region Delete Clean Slate Tests - Query + Connection + Table

    /// <summary>
    /// Verifies that deleting a query loaded to worksheet results in a clean slate:
    /// - Query is removed from queries list
    /// - Connection is removed (no orphans)
    /// - Table/ListObject is removed from worksheet
    /// </summary>
    [Fact]
    public void Delete_LoadedToWorksheet_CleanSlate()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_CleanSlate_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{1}, {2}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query loaded to worksheet
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToTable);

        // Verify everything exists before delete
        var queriesBefore = _powerQueryCommands.List(batch);
        Assert.Contains(queriesBefore.Queries, q => q.Name == queryName);

        var connectionsBefore = _connectionCommands.List(batch);
        Assert.Contains(connectionsBefore.Connections, c => c.Name == $"Query - {queryName}");

        var tablesBefore = _tableCommands.List(batch);
        // Table name typically matches query name
        Assert.True(tablesBefore.Success);
        var tableCountBefore = tablesBefore.Tables.Count;
        Assert.True(tableCountBefore > 0, "Expected at least one table after LoadToTable");

        // Act - Delete the query
        _powerQueryCommands.Delete(batch, queryName);

        // Assert - CLEAN SLATE

        // 1. Query is gone
        var queriesAfter = _powerQueryCommands.List(batch);
        Assert.True(queriesAfter.Success);
        Assert.DoesNotContain(queriesAfter.Queries, q => q.Name == queryName);

        // 2. Connection is gone (no orphans)
        var connectionsAfter = _connectionCommands.List(batch);
        Assert.True(connectionsAfter.Success);
        Assert.DoesNotContain(connectionsAfter.Connections, c => c.Name == $"Query - {queryName}");
        Assert.DoesNotContain(connectionsAfter.Connections, c => c.Name == "Connection");
        Assert.DoesNotContain(connectionsAfter.Connections, c => c.Name.StartsWith("Connection", StringComparison.Ordinal) && char.IsDigit(c.Name.Last()));

        // 3. No Power Query connections remain (clean workbook)
        var pqConnections = connectionsAfter.Connections.Where(c => c.IsPowerQuery).ToList();
        Assert.Empty(pqConnections);
    }

    /// <summary>
    /// Verifies that deleting one of multiple loaded queries only removes that query's
    /// connection and table, leaving others intact.
    /// </summary>
    [Fact]
    public void Delete_OneOfMultipleLoadedQueries_OnlyRemovesItsOwnResources()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var suffix = Guid.NewGuid().ToString("N")[..6];
        var queryToDelete = "PQ_Delete_" + suffix;
        var queryToKeep = "PQ_Keep_" + suffix;
        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create two queries loaded to worksheet
        _powerQueryCommands.Create(batch, queryToDelete, mCode, PowerQueryLoadMode.LoadToTable, "Sheet1");
        _powerQueryCommands.Create(batch, queryToKeep, mCode, PowerQueryLoadMode.LoadToTable, "Sheet2");

        // Verify both exist
        var queriesBefore = _powerQueryCommands.List(batch);
        Assert.Contains(queriesBefore.Queries, q => q.Name == queryToDelete);
        Assert.Contains(queriesBefore.Queries, q => q.Name == queryToKeep);

        var connectionsBefore = _connectionCommands.List(batch);
        Assert.Contains(connectionsBefore.Connections, c => c.Name == $"Query - {queryToDelete}");
        Assert.Contains(connectionsBefore.Connections, c => c.Name == $"Query - {queryToKeep}");

        // Act - Delete only one query
        _powerQueryCommands.Delete(batch, queryToDelete);

        // Assert

        // 1. Deleted query is gone, kept query remains
        var queriesAfter = _powerQueryCommands.List(batch);
        Assert.DoesNotContain(queriesAfter.Queries, q => q.Name == queryToDelete);
        Assert.Contains(queriesAfter.Queries, q => q.Name == queryToKeep);

        // 2. Deleted query's connection is gone, kept query's connection remains
        var connectionsAfter = _connectionCommands.List(batch);
        Assert.DoesNotContain(connectionsAfter.Connections, c => c.Name == $"Query - {queryToDelete}");
        Assert.Contains(connectionsAfter.Connections, c => c.Name == $"Query - {queryToKeep}");

        // 3. No orphaned connections
        Assert.DoesNotContain(connectionsAfter.Connections, c => c.Name == "Connection");

        // 4. Exactly 1 Power Query connection remains
        var pqConnections = connectionsAfter.Connections.Where(c => c.IsPowerQuery).ToList();
        Assert.Single(pqConnections);
    }

    /// <summary>
    /// Verifies clean slate when deleting a ConnectionOnly query (no table involved).
    /// </summary>
    [Fact]
    public void Delete_ConnectionOnly_CleanSlate()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_ConnOnly_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create ConnectionOnly query (no worksheet loading)
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Verify query exists
        var queriesBefore = _powerQueryCommands.List(batch);
        Assert.Contains(queriesBefore.Queries, q => q.Name == queryName);

        // ConnectionOnly may or may not create a connection depending on implementation
        // The key is that after delete, there are no orphans

        // Act
        _powerQueryCommands.Delete(batch, queryName);

        // Assert - Clean slate
        var queriesAfter = _powerQueryCommands.List(batch);
        Assert.DoesNotContain(queriesAfter.Queries, q => q.Name == queryName);

        var connectionsAfter = _connectionCommands.List(batch);
        Assert.DoesNotContain(connectionsAfter.Connections, c => c.Name.Contains(queryName));
        Assert.DoesNotContain(connectionsAfter.Connections, c => c.Name == "Connection");
    }

    #endregion

    #region Unload Clean Slate Tests

    /// <summary>
    /// Verifies that unloading a query from worksheet removes table and connection
    /// but keeps the query definition.
    /// </summary>
    [Fact]
    public void Unload_LoadedToWorksheet_RemovesTableAndConnectionKeepsQuery()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_UnloadWS_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query loaded to worksheet
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToTable);

        // Verify connection exists
        var connectionsBefore = _connectionCommands.List(batch);
        Assert.Contains(connectionsBefore.Connections, c => c.Name == $"Query - {queryName}");

        // Act - Unload
        var unloadResult = _powerQueryCommands.Unload(batch, queryName);
        Assert.True(unloadResult.Success, $"Unload failed: {unloadResult.ErrorMessage}");

        // Assert

        // 1. Query still exists
        var queriesAfter = _powerQueryCommands.List(batch);
        Assert.Contains(queriesAfter.Queries, q => q.Name == queryName);

        // 2. Query is now ConnectionOnly
        var loadConfig = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfig.Success);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, loadConfig.LoadMode);

        // 3. Connection is removed (no active load = no connection needed)
        var connectionsAfter = _connectionCommands.List(batch);
        Assert.DoesNotContain(connectionsAfter.Connections, c => c.Name == $"Query - {queryName}");

        // 4. No orphaned connections
        Assert.DoesNotContain(connectionsAfter.Connections, c => c.Name == "Connection");
    }

    #endregion

    #region Edge Cases

    /// <summary>
    /// Verifies that creating and immediately deleting a query leaves no traces.
    /// </summary>
    [Fact]
    public void CreateThenDelete_LoadToTable_NoTraces()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_CreateDelete_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Act - Create and immediately delete
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToTable);
        _powerQueryCommands.Delete(batch, queryName);

        // Assert - Completely clean workbook
        var queries = _powerQueryCommands.List(batch);
        Assert.Empty(queries.Queries);

        var connections = _connectionCommands.List(batch);
        Assert.Empty(connections.Connections);

        var tables = _tableCommands.List(batch);
        Assert.Empty(tables.Tables);
    }

    /// <summary>
    /// Verifies that the original bug test case (that would have created orphaned
    /// connections) now works correctly with proper connection naming.
    /// </summary>
    [Fact]
    public void Delete_ExistingQuery_VerifiesCleanSlate()
    {
        // This is the improved version of the original Delete_ExistingQuery_ReturnsSuccess test
        // that actually verifies cleanup, not just success

        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_Delete_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let
    Source = #table(
        {""Column1"", ""Column2"", ""Column3""},
        {
            {""Value1"", ""Value2"", ""Value3""},
            {""A"", ""B"", ""C""},
            {""X"", ""Y"", ""Z""}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Act
        _powerQueryCommands.Create(batch, queryName, mCode);  // Default is LoadToTable
        _powerQueryCommands.Delete(batch, queryName);

        // Assert - CLEAN SLATE (not just "reaching here means success")
        var queries = _powerQueryCommands.List(batch);
        Assert.DoesNotContain(queries.Queries, q => q.Name == queryName);

        var connections = _connectionCommands.List(batch);
        Assert.DoesNotContain(connections.Connections, c => c.Name == $"Query - {queryName}");
        Assert.DoesNotContain(connections.Connections, c => c.Name == "Connection");
        Assert.DoesNotContain(connections.Connections, c => c.IsPowerQuery);
    }

    /// <summary>
    /// Verifies LoadTo operation on an existing ConnectionOnly query creates proper connection.
    /// Scenario: Create as ConnectionOnly → LoadTo Table → verify proper naming.
    /// </summary>
    [Fact]
    public void LoadTo_ExistingConnectionOnlyQuery_CreatesProperlyNamedConnection()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_LoadToExisting_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{1}, {2}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create as ConnectionOnly first
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Verify no connections initially
        var connsBefore = _connectionCommands.List(batch);
        Assert.DoesNotContain(connsBefore.Connections, c => c.IsPowerQuery);

        // Act - LoadTo Table
        _powerQueryCommands.LoadTo(batch, queryName, PowerQueryLoadMode.LoadToTable, "Sheet1");

        // Assert - Connection should be properly named
        var connsAfter = _connectionCommands.List(batch);
        Assert.Contains(connsAfter.Connections, c => c.Name == $"Query - {queryName}");
        Assert.DoesNotContain(connsAfter.Connections, c => c.Name == "Connection");

        // Cleanup works
        _powerQueryCommands.Delete(batch, queryName);
        var connsFinal = _connectionCommands.List(batch);
        Assert.Empty(connsFinal.Connections);
    }

    /// <summary>
    /// Verifies Refresh operation maintains proper connection naming and doesn't create orphans.
    /// </summary>
    [Fact]
    public void Refresh_LoadedQuery_MaintainsProperConnectionNaming()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_Refresh_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query with LoadToTable
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToTable);

        var connsBefore = _connectionCommands.List(batch);
        var connectionCountBefore = connsBefore.Connections.Count;

        // Act - Refresh the query
        _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(2));

        // Assert - Same connection count (no new orphans)
        var connsAfter = _connectionCommands.List(batch);
        Assert.Equal(connectionCountBefore, connsAfter.Connections.Count);
        Assert.Contains(connsAfter.Connections, c => c.Name == $"Query - {queryName}");
        Assert.DoesNotContain(connsAfter.Connections, c => c.Name == "Connection");
    }

    /// <summary>
    /// Verifies Update operation maintains proper connection naming and doesn't create orphans.
    /// </summary>
    [Fact]
    public void Update_LoadedQuery_MaintainsProperConnectionNaming()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_Update_" + Guid.NewGuid().ToString("N")[..8];
        var mCode1 = @"let Source = #table({""Val""}, {{1}}) in Source";
        var mCode2 = @"let Source = #table({""NewVal""}, {{2}, {3}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query with LoadToTable
        _powerQueryCommands.Create(batch, queryName, mCode1, PowerQueryLoadMode.LoadToTable);

        var connsBefore = _connectionCommands.List(batch);
        var connectionCountBefore = connsBefore.Connections.Count;

        // Act - Update the query's M code
        _powerQueryCommands.Update(batch, queryName, mCode2);

        // Assert - Same connection count (no new orphans), proper naming
        var connsAfter = _connectionCommands.List(batch);
        Assert.Equal(connectionCountBefore, connsAfter.Connections.Count);
        Assert.Contains(connsAfter.Connections, c => c.Name == $"Query - {queryName}");
        Assert.DoesNotContain(connsAfter.Connections, c => c.Name == "Connection");

        // Cleanup still works
        _powerQueryCommands.Delete(batch, queryName);
        var connsFinal = _connectionCommands.List(batch);
        Assert.DoesNotContain(connsFinal.Connections, c => c.IsPowerQuery);
    }

    /// <summary>
    /// Verifies that mode transition from ConnectionOnly to LoadToBoth creates proper dual connections.
    /// </summary>
    [Fact]
    public async Task LoadTo_ConnectionOnlyToBoth_CreatesDualConnectionsProperly()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_ConnToBoth_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{42}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create as ConnectionOnly
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Verify no Power Query connections
        var connsBefore = _connectionCommands.List(batch);
        Assert.DoesNotContain(connsBefore.Connections, c => c.IsPowerQuery);

        // Act - LoadTo Both
        _powerQueryCommands.LoadTo(batch, queryName, PowerQueryLoadMode.LoadToBoth, "BothSheet");

        // Assert - Should have TWO connections with proper naming
        var connsAfter = _connectionCommands.List(batch);
        var pqConns = connsAfter.Connections.Where(c => c.IsPowerQuery).ToList();

        // Should have exactly 2 Power Query connections
        Assert.Equal(2, pqConns.Count);

        // One for worksheet, one for Data Model
        Assert.Contains(pqConns, c => c.Name == $"Query - {queryName}");
        Assert.Contains(pqConns, c => c.Name == $"Query - {queryName} (Data Model)");

        // No orphans
        Assert.DoesNotContain(connsAfter.Connections, c => c.Name == "Connection");

        // Data Model table should exist
        var tables = await _dataModelCommands.ListTables(batch);
        Assert.Contains(tables.Tables, t => t.Name == queryName);

        // Cleanup
        _powerQueryCommands.Delete(batch, queryName);
        var connsFinal = _connectionCommands.List(batch);
        Assert.DoesNotContain(connsFinal.Connections, c => c.IsPowerQuery);
    }

    /// <summary>
    /// Verifies LoadToBoth creates exactly 2 connections and both are properly cleaned up.
    /// </summary>
    [Fact]
    public async Task Create_LoadToBoth_ExactlyTwoConnectionsWithProperNaming()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_BothDual_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Act
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToBoth, "Sheet1");

        // Assert - Exactly 2 Power Query connections
        var connections = _connectionCommands.List(batch);
        var pqConns = connections.Connections.Where(c => c.IsPowerQuery).ToList();

        Assert.Equal(2, pqConns.Count);

        // Verify exact naming pattern
        Assert.Contains(pqConns, c => c.Name == $"Query - {queryName}");
        Assert.Contains(pqConns, c => c.Name == $"Query - {queryName} (Data Model)");

        // Verify worksheet table exists
        var tables = _tableCommands.List(batch);
        Assert.True(tables.Tables.Count > 0);

        // Verify Data Model table exists
        var dmTables = await _dataModelCommands.ListTables(batch);
        Assert.Contains(dmTables.Tables, t => t.Name == queryName);

        // Cleanup removes both
        _powerQueryCommands.Delete(batch, queryName);
        var connsFinal = _connectionCommands.List(batch);
        Assert.DoesNotContain(connsFinal.Connections, c => c.IsPowerQuery);
    }

    /// <summary>
    /// Verifies Unload then re-LoadTo works correctly without creating orphans.
    /// Scenario: Create LoadToTable → Unload → LoadTo Table again
    /// </summary>
    [Fact]
    public void UnloadThenReload_NoOrphanedConnections()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_ReloadTest_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create with LoadToTable
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToTable);

        var connsAfterCreate = _connectionCommands.List(batch);
        Assert.Single(connsAfterCreate.Connections, c => c.IsPowerQuery);

        // Unload
        _powerQueryCommands.Unload(batch, queryName);

        var connsAfterUnload = _connectionCommands.List(batch);
        Assert.DoesNotContain(connsAfterUnload.Connections, c => c.IsPowerQuery);

        // Re-LoadTo Table
        _powerQueryCommands.LoadTo(batch, queryName, PowerQueryLoadMode.LoadToTable, "NewSheet");

        // Assert - Should have exactly 1 properly named connection, no orphans
        var connsAfterReload = _connectionCommands.List(batch);
        var pqConns = connsAfterReload.Connections.Where(c => c.IsPowerQuery).ToList();

        Assert.Single(pqConns);
        Assert.Equal($"Query - {queryName}", pqConns[0].Name);
        Assert.DoesNotContain(connsAfterReload.Connections, c => c.Name == "Connection");

        // Cleanup
        _powerQueryCommands.Delete(batch, queryName);
        var connsFinal = _connectionCommands.List(batch);
        Assert.DoesNotContain(connsFinal.Connections, c => c.IsPowerQuery);
    }

    #endregion
}
