using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePowerQueryLifecycleTests
{
    [Fact]
    public void Unload_DataModelOnly_RemovesDataModelConnection()
    {
        var queryName = UniqueCleanupName("PQ_UnloadDM");
        CreateCleanupQuery(queryName, PowerQueryLoadMode.LoadToDataModel);

        var tablesBefore = _dataModel.ListTables(_fixture.BatchToken);
        Assert.True(tablesBefore.Success);
        Assert.Contains(tablesBefore.Tables, table => table.Name == queryName);
        var connectionsBefore = _connections.List(_fixture.BatchToken);
        Assert.True(connectionsBefore.Success);
        Assert.Contains(
            connectionsBefore.Connections,
            connection => connection.Name.Contains($"Query - {queryName}"));

        var unloadResult = _queries.Unload(_fixture.BatchToken, queryName);

        Assert.True(unloadResult.Success, $"Unload failed: {unloadResult.ErrorMessage}");
        var connectionsAfter = _connections.List(_fixture.BatchToken);
        Assert.True(connectionsAfter.Success);
        Assert.DoesNotContain(
            connectionsAfter.Connections,
            connection => connection.Name.Contains($"Query - {queryName}"));
        var queries = _queries.List(_fixture.BatchToken);
        Assert.True(queries.Success);
        Assert.Contains(queries.Queries, query => query.Name == queryName);
        var loadConfig = _queries.GetLoadConfig(_fixture.BatchToken, queryName);
        Assert.True(loadConfig.Success);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, loadConfig.LoadMode);
    }

    [Fact]
    public void Unload_LoadToBoth_RemovesBothWorksheetAndDataModelConnection()
    {
        var queryName = UniqueCleanupName("PQ_UnloadBoth");
        var sheetName = UniqueCleanupName("BothSheet");
        CreateCleanupQuery(queryName, PowerQueryLoadMode.LoadToBoth, sheetName);

        var tablesBefore = _dataModel.ListTables(_fixture.BatchToken);
        Assert.True(tablesBefore.Success);
        Assert.Contains(tablesBefore.Tables, table => table.Name == queryName);
        var connectionsBefore = _connections.List(_fixture.BatchToken);
        Assert.True(connectionsBefore.Success);
        Assert.Contains(
            connectionsBefore.Connections,
            connection => connection.Name.Contains($"Query - {queryName}"));

        var unloadResult = _queries.Unload(_fixture.BatchToken, queryName);

        Assert.True(unloadResult.Success, $"Unload failed: {unloadResult.ErrorMessage}");
        var connectionsAfter = _connections.List(_fixture.BatchToken);
        Assert.True(connectionsAfter.Success);
        Assert.DoesNotContain(
            connectionsAfter.Connections,
            connection => connection.Name.Contains($"Query - {queryName}"));
        var loadConfig = _queries.GetLoadConfig(_fixture.BatchToken, queryName);
        Assert.True(loadConfig.Success);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, loadConfig.LoadMode);
    }

    [Fact]
    public void Delete_DataModelOnly_RemovesDataModelConnection()
    {
        var queryName = UniqueCleanupName("PQ_DeleteDM");
        CreateCleanupQuery(queryName, PowerQueryLoadMode.LoadToDataModel);

        var tablesBefore = _dataModel.ListTables(_fixture.BatchToken);
        Assert.True(tablesBefore.Success);
        Assert.Contains(tablesBefore.Tables, table => table.Name == queryName);
        var connectionsBefore = _connections.List(_fixture.BatchToken);
        Assert.True(connectionsBefore.Success);
        Assert.Contains(
            connectionsBefore.Connections,
            connection => connection.Name.Contains($"Query - {queryName}"));

        _queries.Delete(_fixture.BatchToken, queryName);
        _fixture.ForgetPowerQuery(queryName);

        var queries = _queries.List(_fixture.BatchToken);
        Assert.True(queries.Success);
        Assert.DoesNotContain(queries.Queries, query => query.Name == queryName);
        var connectionsAfter = _connections.List(_fixture.BatchToken);
        Assert.True(connectionsAfter.Success);
        Assert.DoesNotContain(
            connectionsAfter.Connections,
            connection => connection.Name.Contains($"Query - {queryName}"));
    }

    [Fact]
    public void Delete_LoadToBoth_RemovesBothWorksheetAndDataModelConnection()
    {
        var queryName = UniqueCleanupName("PQ_DeleteBoth");
        var sheetName = UniqueCleanupName("DeleteBothSheet");
        CreateCleanupQuery(queryName, PowerQueryLoadMode.LoadToBoth, sheetName);

        var tablesBefore = _dataModel.ListTables(_fixture.BatchToken);
        Assert.True(tablesBefore.Success);
        Assert.Contains(tablesBefore.Tables, table => table.Name == queryName);
        var connectionsBefore = _connections.List(_fixture.BatchToken);
        Assert.True(connectionsBefore.Success);
        Assert.Contains(
            connectionsBefore.Connections,
            connection => connection.Name.Contains($"Query - {queryName}"));

        _queries.Delete(_fixture.BatchToken, queryName);
        _fixture.ForgetPowerQuery(queryName);

        var queries = _queries.List(_fixture.BatchToken);
        Assert.True(queries.Success);
        Assert.DoesNotContain(queries.Queries, query => query.Name == queryName);
        var connectionsAfter = _connections.List(_fixture.BatchToken);
        Assert.True(connectionsAfter.Success);
        Assert.DoesNotContain(
            connectionsAfter.Connections,
            connection => connection.Name.Contains($"Query - {queryName}"));
    }

    private void CreateCleanupQuery(
        string queryName,
        PowerQueryLoadMode loadMode,
        string? sheetName = null)
    {
        _queries.Create(
            _fixture.BatchToken,
            queryName,
            "let Source = #table({\"Val\"}, {{1}}) in Source",
            loadMode,
            sheetName);
        _fixture.RegisterPowerQueryForCleanup(queryName);
        if (sheetName is not null)
        {
            _fixture.RegisterSheetForCleanup(sheetName);
        }
    }

    private static string UniqueCleanupName(string prefix) =>
        $"{prefix}_{Guid.NewGuid():N}"[..Math.Min(prefix.Length + 9, 31)];
}
