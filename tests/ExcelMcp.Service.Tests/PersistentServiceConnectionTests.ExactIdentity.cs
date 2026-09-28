using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    [Fact]
    public void Delete_Connection_PreservesUnrelatedConnectionDataQueryTable()
    {
        var batch = _fixture.BatchToken;
        var connectionName = ExactIdentityName("Connection");
        var queryTableName = $"{connectionName}Data";
        var sheetName = _fixture.CreateTestSheet(batch);
        var sourceFile = CreateTextSource();
        var queryTables = _fixture.CreateCommands<IQueryTableCommands>();
        _connections.Create(
            batch,
            connectionName,
            "ODBC;DSN=DeleteIdentityTest");
        _fixture.RegisterConnectionForCleanup(connectionName);
        queryTables.CreateText(
            batch,
            queryTableName,
            sourceFile,
            sheetName,
            "A1");

        var result = _connections.Delete(batch, connectionName);
        _fixture.ForgetConnection(connectionName);

        Assert.True(result.Success);
        Assert.Contains(
            queryTables.List(batch).QueryTables,
            queryTable => queryTable.Name == queryTableName);
    }

    [Fact]
    public void LoadTo_Connection_PreservesUnrelatedConnectionDataQueryTable()
    {
        var batch = _fixture.BatchToken;
        var connectionName = ExactIdentityName("Connection");
        var queryTableName = $"{connectionName}Data";
        var textSheetName = _fixture.CreateTestSheet(batch);
        var loadSheetName = UniqueSheetName("ConnectionLoad");
        var textSourceFile = CreateTextSource();
        var aceSourceFile = Path.Combine(
            Path.GetDirectoryName(_fixture.WorkbookPath)!,
            $"ACE_{Guid.NewGuid():N}.xlsx");
        var queryTables = _fixture.CreateCommands<IQueryTableCommands>();
        Sbroenne.ExcelMcp.Core.Tests.Helpers.AceOleDbTestHelper
            .CreateExcelDataSource(aceSourceFile);
        _connections.Create(
            batch,
            connectionName,
            Sbroenne.ExcelMcp.Core.Tests.Helpers.AceOleDbTestHelper
                .GetExcelConnectionString(aceSourceFile),
            commandText: Sbroenne.ExcelMcp.Core.Tests.Helpers.AceOleDbTestHelper
                .GetDefaultCommandText());
        _fixture.RegisterConnectionForCleanup(connectionName);
        queryTables.CreateText(
            batch,
            queryTableName,
            textSourceFile,
            textSheetName,
            "A1");

        try
        {
            var result = _connections.LoadTo(
                batch,
                connectionName,
                loadSheetName);
            _fixture.RegisterSheetForCleanup(loadSheetName);

            Assert.True(result.Success);
            var loaded = queryTables.List(batch);
            Assert.Contains(
                loaded.QueryTables,
                queryTable => queryTable.Name == queryTableName);
            Assert.Contains(
                loaded.QueryTables,
                queryTable => queryTable.Name == connectionName);

            _connections.Delete(batch, connectionName);
            _fixture.ForgetConnection(connectionName);

            var afterDelete = queryTables.List(batch);
            Assert.Contains(
                afterDelete.QueryTables,
                queryTable => queryTable.Name == queryTableName);
            Assert.DoesNotContain(
                afterDelete.QueryTables,
                queryTable => queryTable.Name == connectionName);
        }
        finally
        {
            File.Delete(aceSourceFile);
        }
    }

    [Fact]
    public void Delete_ExactConnection_PreservesPowerQueryAlias()
    {
        var batch = _fixture.BatchToken;
        var name = ExactIdentityName("Connection");
        var sheetName = _fixture.CreateTestSheet(batch);
        var powerQueries = _fixture.CreateCommands<IPowerQueryCommands>();
        const string mCode =
            "let Source = #table({\"Value\"}, {{1}}) in Source";
        powerQueries.Create(
            batch,
            name,
            mCode,
            PowerQueryLoadMode.LoadToTable,
            sheetName);
        _fixture.RegisterPowerQueryForCleanup(name);
        _connections.Create(
            batch,
            name,
            "ODBC;DSN=ExactConnectionTest");
        _fixture.RegisterConnectionForCleanup(name);

        var result = _connections.Delete(batch, name);
        _fixture.ForgetConnection(name);

        Assert.True(result.Success);
        Assert.Contains(
            _connections.List(batch).Connections,
            connection => connection.Name == $"Query - {name}");
        Assert.Equal(
            PowerQueryLoadMode.LoadToTable,
            powerQueries.GetLoadConfig(batch, name).LoadMode);
    }

    private string CreateTextSource() =>
        _fixture.CreateInputFile(
            ".csv",
            "Name,Value\r\nPreserved,1\r\n");

    private static string ExactIdentityName(string prefix) =>
        $"{prefix}_{Guid.NewGuid():N}"[..24];
}
