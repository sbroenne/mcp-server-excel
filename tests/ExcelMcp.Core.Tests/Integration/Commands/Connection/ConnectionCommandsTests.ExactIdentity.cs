using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

public partial class ConnectionCommandsTests
{
    [Fact]
    public void Delete_Connection_PreservesUnrelatedConnectionDataQueryTable()
    {
        var testFile = _fixture.CreateTestFile();
        var sourceFile = CreateTextSource();
        var queryTableCommands = new QueryTableCommands();

        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Create(batch, "Connection", "ODBC;DSN=DeleteIdentityTest");
        queryTableCommands.CreateText(
            batch,
            "ConnectionData",
            sourceFile,
            "Sheet1",
            "A1");

        var result = _commands.Delete(batch, "Connection");

        Assert.True(result.Success);
        Assert.Contains(
            queryTableCommands.List(batch).QueryTables,
            queryTable => queryTable.Name == "ConnectionData");
    }

    [Fact]
    public void LoadTo_Connection_PreservesUnrelatedConnectionDataQueryTable()
    {
        var testFile = _fixture.CreateTestFile();
        var textSourceFile = CreateTextSource();
        var aceSourceFile = Path.Join(_fixture.TempDir, $"ACE_{Guid.NewGuid():N}.xlsx");
        var queryTableCommands = new QueryTableCommands();

        AceOleDbTestHelper.CreateExcelDataSource(aceSourceFile);
        ConnectionTestHelper.CreateAceOleDbConnection(testFile, "Connection", aceSourceFile);

        using var batch = ExcelSession.BeginBatch(testFile);
        queryTableCommands.CreateText(
            batch,
            "ConnectionData",
            textSourceFile,
            "Sheet1",
            "A1");

        var result = _commands.LoadTo(batch, "Connection", "ConnectionLoad");

        Assert.True(result.Success);
        var queryTables = queryTableCommands.List(batch);
        Assert.Contains(queryTables.QueryTables, queryTable => queryTable.Name == "ConnectionData");
        Assert.Contains(queryTables.QueryTables, queryTable => queryTable.Name == "Connection");

        _commands.Delete(batch, "Connection");

        var queryTablesAfterDelete = queryTableCommands.List(batch);
        Assert.Contains(
            queryTablesAfterDelete.QueryTables,
            queryTable => queryTable.Name == "ConnectionData");
        Assert.DoesNotContain(
            queryTablesAfterDelete.QueryTables,
            queryTable => queryTable.Name == "Connection");
    }

    [Fact]
    public void Delete_ExactConnection_PreservesPowerQueryAlias()
    {
        var testFile = _fixture.CreateTestFile();
        var powerQueryCommands = new PowerQueryCommands(new DataModelCommands());
        const string mCode = "let Source = #table({\"Value\"}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testFile);
        powerQueryCommands.Create(
            batch,
            "Connection",
            mCode,
            PowerQueryLoadMode.LoadToTable,
            "QueryData");
        _commands.Create(batch, "Connection", "ODBC;DSN=ExactConnectionTest");

        var result = _commands.Delete(batch, "Connection");

        Assert.True(result.Success);
        Assert.Contains(
            _commands.List(batch).Connections,
            connection => connection.Name == "Query - Connection");
        Assert.Equal(
            PowerQueryLoadMode.LoadToTable,
            powerQueryCommands.GetLoadConfig(batch, "Connection").LoadMode);
    }

    private string CreateTextSource()
    {
        var sourceFile = Path.Join(_fixture.TempDir, $"ConnectionData_{Guid.NewGuid():N}.csv");
        System.IO.File.WriteAllText(sourceFile, "Name,Value\r\nPreserved,1\r\n");
        return sourceFile;
    }
}
