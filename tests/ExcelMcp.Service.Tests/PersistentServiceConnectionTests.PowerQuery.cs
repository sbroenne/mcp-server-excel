using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    [Fact]
    public void Delete_OrphanedPowerQueryConnection_GenericName_Succeeds()
    {
        const string connectionName = "Connection";
        AddMashupConnection(connectionName, $"Missing_{Guid.NewGuid():N}");
        _fixture.RegisterConnectionForCleanup(connectionName);

        var before = _connections.List(_fixture.BatchToken);
        var orphaned = Assert.Single(
            before.Connections,
            connection => connection.Name == connectionName);
        Assert.True(orphaned.IsPowerQuery);

        var result = _connections.Delete(
            _fixture.BatchToken,
            connectionName);
        _fixture.ForgetConnection(connectionName);

        Assert.True(result.Success);
        Assert.DoesNotContain(
            _connections.List(_fixture.BatchToken).Connections,
            connection => connection.Name == connectionName);
    }

    [Fact]
    public void Delete_OrphanedPowerQueryConnection_StandardNameMissingQuery_Succeeds()
    {
        var missingQueryName = $"Missing_{Guid.NewGuid():N}"[..24];
        var connectionName = $"Query - {missingQueryName}";
        AddMashupConnection(connectionName, missingQueryName);
        _fixture.RegisterConnectionForCleanup(connectionName);

        var before = _connections.List(_fixture.BatchToken);
        var orphaned = Assert.Single(
            before.Connections,
            connection => connection.Name == connectionName);
        Assert.True(orphaned.IsPowerQuery);

        var result = _connections.Delete(
            _fixture.BatchToken,
            connectionName);
        _fixture.ForgetConnection(connectionName);

        Assert.True(result.Success);
        Assert.DoesNotContain(
            _connections.List(_fixture.BatchToken).Connections,
            connection => connection.Name == connectionName);
    }

    [Fact]
    public void Delete_ValidPowerQueryConnection_ThrowsWithRedirect()
    {
        var queryName = $"Valid_{Guid.NewGuid():N}"[..24];
        var powerQueries = _fixture.CreateCommands<IPowerQueryCommands>();
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        powerQueries.Create(
            _fixture.BatchToken,
            queryName,
            "let Source = #table({\"Value\"}, {{1}}) in Source",
            PowerQueryLoadMode.LoadToTable,
            sheetName);
        _fixture.RegisterPowerQueryForCleanup(queryName);
        var connectionName = $"Query - {queryName}";

        var before = _connections.List(_fixture.BatchToken);
        var validConnection = Assert.Single(
            before.Connections,
            connection => connection.Name == connectionName);
        Assert.True(validConnection.IsPowerQuery);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _connections.Delete(
                _fixture.BatchToken,
                connectionName));

        Assert.Contains(
            "powerquery",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    private void AddMashupConnection(
        string connectionName,
        string location) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic? connections = null;
            dynamic? connection = null;
            try
            {
                connections = ctx.Book.Connections;
                connection = connections.Add2(
                    Name: connectionName,
                    Description: "Orphaned Power Query connection test",
                    ConnectionString:
                        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;" +
                        $"Data Source=$Workbook$;Location={location};" +
                        "Extended Properties=\"\"",
                    CommandText: $"SELECT * FROM [{location}]",
                    lCmdtype: 2,
                    CreateModelConnection: false,
                    ImportRelationships: false);
            }
            finally
            {
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref connections);
            }
        });
}
