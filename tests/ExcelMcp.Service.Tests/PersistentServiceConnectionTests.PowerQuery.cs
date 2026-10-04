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
        var retained = SeedRetainedConnection();
        AddMashupConnection(connectionName, $"Missing_{Guid.NewGuid():N}");
        _fixture.RegisterConnectionForCleanup(connectionName);

        var before = RequireSuccess(_connections.List(_fixture.BatchToken));
        var orphaned = Assert.Single(
            before.Connections,
            connection => connection.Name == connectionName);
        Assert.True(orphaned.IsPowerQuery);

        var result = _connections.Delete(
            _fixture.BatchToken,
            connectionName);
        _fixture.ForgetConnection(connectionName);

        Assert.True(result.Success);
        RequireSuccess(result);
        Assert.DoesNotContain(
            RequireSuccess(_connections.List(_fixture.BatchToken)).Connections,
            connection => connection.Name == connectionName);
        AssertRetainedConnection(retained);
    }

    [Fact]
    public void Delete_OrphanedPowerQueryConnection_StandardNameMissingQuery_Succeeds()
    {
        var missingQueryName = $"Missing_{Guid.NewGuid():N}"[..24];
        var connectionName = $"Query - {missingQueryName}";
        var retained = SeedRetainedConnection();
        AddMashupConnection(connectionName, missingQueryName);
        _fixture.RegisterConnectionForCleanup(connectionName);

        var before = RequireSuccess(_connections.List(_fixture.BatchToken));
        var orphaned = Assert.Single(
            before.Connections,
            connection => connection.Name == connectionName);
        Assert.True(orphaned.IsPowerQuery);

        var result = _connections.Delete(
            _fixture.BatchToken,
            connectionName);
        _fixture.ForgetConnection(connectionName);

        Assert.True(result.Success);
        RequireSuccess(result);
        Assert.DoesNotContain(
            RequireSuccess(_connections.List(_fixture.BatchToken)).Connections,
            connection => connection.Name == connectionName);
        AssertRetainedConnection(retained);
    }

    [Fact]
    public void Delete_ValidPowerQueryConnection_ThrowsWithRedirect()
    {
        var queryName = $"Valid_{Guid.NewGuid():N}"[..24];
        var powerQueries = _fixture.CreateCommands<IPowerQueryCommands>();
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(powerQueries.Create(
            _fixture.BatchToken,
            queryName,
            "let Source = #table({\"Value\"}, {{1}}) in Source",
            PowerQueryLoadMode.LoadToTable,
            sheetName));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        var connectionName = $"Query - {queryName}";

        var before = RequireSuccess(_connections.List(_fixture.BatchToken));
        var validConnection = Assert.Single(
            before.Connections,
            connection => connection.Name == connectionName);
        Assert.True(validConnection.IsPowerQuery);
        var nativeBefore = ReadNativeConnection(connectionName);
        var cellsBefore = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1:A2")).Values;
        Assert.Equal("Value", Assert.Single(cellsBefore[0]));
        Assert.Equal(1, Convert.ToInt32(Assert.Single(cellsBefore[1]), System.Globalization.CultureInfo.InvariantCulture));

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _connections.Delete(
                _fixture.BatchToken,
                connectionName));

        Assert.Contains(
            "powerquery",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        Assert.Equal(nativeBefore, ReadNativeConnection(connectionName));
        Assert.Equal(PowerQueryLoadMode.LoadToTable,
            RequireSuccess(powerQueries.GetLoadConfig(_fixture.BatchToken, queryName)).LoadMode);
        var cellsAfter = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1:A2")).Values;
        Assert.Equal(cellsBefore.Count, cellsAfter.Count);
        for (var index = 0; index < cellsBefore.Count; index++)
        {
            Assert.Equal(cellsBefore[index], cellsAfter[index]);
        }
    }

    private void AddMashupConnection(
        string connectionName,
        string location) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Microsoft.Office.Interop.Excel.Connections? connections = null;
            Microsoft.Office.Interop.Excel.WorkbookConnection? connection = null;
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
