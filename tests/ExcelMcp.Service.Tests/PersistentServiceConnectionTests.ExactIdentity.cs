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
        RequireSuccess(_connections.Create(
            batch,
            connectionName,
            "ODBC;DSN=DeleteIdentityTest"));
        _fixture.RegisterConnectionForCleanup(connectionName);
        RequireSuccess(queryTables.CreateText(
            batch,
            queryTableName,
            sourceFile,
            sheetName,
            "A1"));
        AssertPreservedTextData(sheetName);
        var textBefore = RequireSuccess(queryTables.View(batch, sheetName, queryTableName));

        var result = _connections.Delete(batch, connectionName);
        _fixture.ForgetConnection(connectionName);

        Assert.True(result.Success);
        RequireSuccess(result);
        Assert.Contains(
            RequireSuccess(queryTables.List(batch)).QueryTables,
            queryTable => queryTable.Name == queryTableName);
        AssertPreservedTextData(sheetName);
        Assert.Equal(System.Text.Json.JsonSerializer.Serialize(textBefore),
            System.Text.Json.JsonSerializer.Serialize(RequireSuccess(queryTables.View(batch, sheetName, queryTableName))));
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
        RequireSuccess(_connections.Create(
            batch,
            connectionName,
            Sbroenne.ExcelMcp.Core.Tests.Helpers.AceOleDbTestHelper
                .GetExcelConnectionString(aceSourceFile),
            commandText: Sbroenne.ExcelMcp.Core.Tests.Helpers.AceOleDbTestHelper
                .GetDefaultCommandText()));
        _fixture.RegisterConnectionForCleanup(connectionName);
        RequireSuccess(queryTables.CreateText(
            batch,
            queryTableName,
            textSourceFile,
            textSheetName,
            "A1"));
        AssertPreservedTextData(textSheetName);
        var textBefore = RequireSuccess(queryTables.View(batch, textSheetName, queryTableName));

        try
        {
            var result = _connections.LoadTo(
                batch,
                connectionName,
                loadSheetName);
            _fixture.RegisterSheetForCleanup(loadSheetName);

            Assert.True(result.Success);
            RequireSuccess(result);
            var loaded = RequireSuccess(queryTables.List(batch));
            Assert.Contains(
                loaded.QueryTables,
                queryTable => queryTable.Name == queryTableName);
            Assert.Contains(
                loaded.QueryTables,
                queryTable => queryTable.Name == connectionName);
            AssertAceData(loadSheetName, 19.99);
            AssertPreservedTextData(textSheetName);
            Assert.Equal(System.Text.Json.JsonSerializer.Serialize(textBefore),
                System.Text.Json.JsonSerializer.Serialize(RequireSuccess(queryTables.View(batch, textSheetName, queryTableName))));

            RequireSuccess(_connections.Delete(batch, connectionName));
            _fixture.ForgetConnection(connectionName);

            var afterDelete = RequireSuccess(queryTables.List(batch));
            Assert.Contains(
                afterDelete.QueryTables,
                queryTable => queryTable.Name == queryTableName);
            Assert.DoesNotContain(
                afterDelete.QueryTables,
                queryTable => queryTable.Name == connectionName);
            AssertPreservedTextData(textSheetName);
            Assert.Equal(System.Text.Json.JsonSerializer.Serialize(textBefore),
                System.Text.Json.JsonSerializer.Serialize(RequireSuccess(queryTables.View(batch, textSheetName, queryTableName))));
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
        RequireSuccess(powerQueries.Create(
            batch,
            name,
            mCode,
            PowerQueryLoadMode.LoadToTable,
            sheetName));
        _fixture.RegisterPowerQueryForCleanup(name);
        RequireSuccess(_connections.Create(
            batch,
            name,
            "ODBC;DSN=ExactConnectionTest"));
        _fixture.RegisterConnectionForCleanup(name);
        var queryBefore = RequireSuccess(powerQueries.View(batch, name));
        var sourceBefore = ReadNativeConnection($"Query - {name}");

        var result = _connections.Delete(batch, name);
        _fixture.ForgetConnection(name);

        Assert.True(result.Success);
        RequireSuccess(result);
        Assert.Contains(
            RequireSuccess(_connections.List(batch)).Connections,
            connection => connection.Name == $"Query - {name}");
        Assert.Equal(
            PowerQueryLoadMode.LoadToTable,
            RequireSuccess(powerQueries.GetLoadConfig(batch, name)).LoadMode);
        Assert.Equal(sourceBefore, ReadNativeConnection($"Query - {name}"));
        Assert.Equal(System.Text.Json.JsonSerializer.Serialize(queryBefore),
            System.Text.Json.JsonSerializer.Serialize(RequireSuccess(powerQueries.View(batch, name))));
        var values = RequireSuccess(_commands.GetValues(batch, sheetName, "A1:A2")).Values;
        Assert.Equal("Value", Assert.Single(values[0]));
        Assert.Equal(1, Convert.ToInt32(Assert.Single(values[1]), System.Globalization.CultureInfo.InvariantCulture));
    }

    private string CreateTextSource() =>
        _fixture.CreateInputFile(
            ".csv",
            "Name,Value\r\nPreserved,1\r\n");

    private static string ExactIdentityName(string prefix) =>
        $"{prefix}_{Guid.NewGuid():N}"[..24];
}
