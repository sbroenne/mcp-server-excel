using System.Data.Common;
using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePowerQueryExactIdentityTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private const string PrefixQueryMCode =
        "let Source = #table({\"Value\"}, {{1}}) in Source";
    private const string NeighborQueryMCode =
        "let Source = #table({\"Value\"}, {{23}, {47}}) in Source";

    private readonly IPowerQueryCommands _queries =
        fixture.CreateCommands<IPowerQueryCommands>();
    private readonly IConnectionCommands _connections =
        fixture.CreateCommands<IConnectionCommands>();
    private readonly ITableCommands _tables =
        fixture.CreateCommands<ITableCommands>();
    private readonly IDataModelCommands _dataModel =
        fixture.CreateCommands<IDataModelCommands>();

    [Fact]
    public async Task Refresh_MissingQuery_ReturnsCategorizedNotFound()
    {
        CreatePrefixQueries(PowerQueryLoadMode.LoadToTable);
        var response = await _fixture.SendForFailureAsync(
            "powerquery.refresh",
            new { queryName = "MissingQuery", timeout = 30 });

        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal("NotFound", response.ErrorCategory);
        Assert.Contains("Query 'MissingQuery' not found.", response.ErrorMessage);
        AssertPrefixStored();
        AssertWorksheetLoadPreserved("AA");
    }

    private void AssertGenericConnection() =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Connections? connections = null;
            Excel.WorkbookConnection? connection = null;
            Excel.ODBCConnection? odbc = null;
            try
            {
                connections = context.Book.Connections;
                connection = connections.Item("Connection");
                Assert.Equal("Connection", connection.Name);
                Assert.Equal(Excel.XlConnectionType.xlConnectionTypeODBC, connection.Type);
                odbc = connection.ODBCConnection;
                string text = Convert.ToString(odbc.Connection, CultureInfo.InvariantCulture) ?? "";
                if (text.StartsWith("ODBC;", StringComparison.OrdinalIgnoreCase)) { text = text[5..]; }
                var properties = new DbConnectionStringBuilder { ConnectionString = text };
                Assert.Equal("PreservedGenericConnection", properties["DSN"]);
                Assert.False(odbc.Refreshing);
            }
            finally
            {
                ComUtilities.Release(ref odbc);
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref connections);
            }
        });

    [Fact]
    public async Task ExactIdentity_ReadAndRefreshPaths_DoNotTreatAAAsA()
    {
        CreatePrefixQueries(PowerQueryLoadMode.LoadToTable);

        var listResult = RequireSuccess(_queries.List(_fixture.BatchToken));
        Assert.True(listResult.Queries.Single(query => query.Name == "A").IsConnectionOnly);
        Assert.False(listResult.Queries.Single(query => query.Name == "AA").IsConnectionOnly);

        var viewResult = RequireSuccess(_queries.View(_fixture.BatchToken, "a"));
        Assert.Equal(PrefixQueryMCode, viewResult.MCode);
        Assert.True(viewResult.IsConnectionOnly);

        var loadConfig = RequireSuccess(_queries.GetLoadConfig(_fixture.BatchToken, "a"));
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, loadConfig.LoadMode);

        var response = await _fixture.SendForFailureAsync(
            "powerquery.refresh",
            new { queryName = "a", timeout = 30 });
        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal("Prerequisite", response.ErrorCategory);
        Assert.Contains(
            "Could not find connection or table for query 'a'",
            response.ErrorMessage);

        AssertPrefixStored();
        AssertWorksheetLoadPreserved("AA");
    }

    [Fact]
    public void LoadTo_PrefixQuery_PreservesAAWorksheetDestination()
    {
        CreatePrefixQueries(PowerQueryLoadMode.LoadToTable);

        var result = RequireSuccess(_queries.LoadTo(
            _fixture.BatchToken,
            "A",
            PowerQueryLoadMode.LoadToTable,
            "AData",
            "A1"));
        _fixture.RegisterSheetForCleanup("AData");

        Assert.True(result.Success);
        Assert.Equal(
            PowerQueryLoadMode.LoadToTable,
            RequireSuccess(_queries.GetLoadConfig(_fixture.BatchToken, "A")).LoadMode);
        PowerQueryStateAssertions.AssertStored(_fixture, "A", PrefixQueryMCode,
            PowerQueryLoadMode.LoadToTable, "AData", ["Value"], [[1]]);
        AssertWorksheetLoadPreserved("AA");
    }

    [Fact]
    public void Unload_PrefixQuery_PreservesAAWorksheetDestination()
    {
        CreatePrefixQueries(PowerQueryLoadMode.LoadToTable);

        var result = RequireSuccess(_queries.Unload(_fixture.BatchToken, "A"));

        Assert.True(result.Success);
        AssertPrefixStored();
        AssertWorksheetLoadPreserved("AA");
    }

    [Fact]
    public void Delete_PrefixQuery_PreservesAAWorksheetDestination()
    {
        CreatePrefixQueries(PowerQueryLoadMode.LoadToTable);

        var result = RequireSuccess(_queries.Delete(_fixture.BatchToken, "a"));
        _fixture.ForgetPowerQuery("A");

        Assert.True(result.Success);
        Assert.DoesNotContain(
            RequireSuccess(_queries.List(_fixture.BatchToken)).Queries,
            query => query.Name == "A");
        AssertWorksheetLoadPreserved("AA");
    }

    [Fact]
    public void Unload_PrefixQuery_PreservesAADataModelDestination()
    {
        CreatePrefixQueries(PowerQueryLoadMode.LoadToDataModel);

        var result = RequireSuccess(_queries.Unload(_fixture.BatchToken, "A"));

        Assert.True(result.Success);
        var tables = RequireSuccess(_dataModel.ListTables(_fixture.BatchToken));
        Assert.Contains(tables.Tables, table => table.Name == "AA");
        Assert.Equal(
            PowerQueryLoadMode.LoadToDataModel,
            RequireSuccess(_queries.GetLoadConfig(_fixture.BatchToken, "AA")).LoadMode);
        AssertPrefixStored();
        PowerQueryStateAssertions.AssertStored(_fixture, "AA", NeighborQueryMCode,
            PowerQueryLoadMode.LoadToDataModel, null, ["Value"], [[23], [47]]);
    }

    [Fact]
    public void ConnectionOnlyQuery_DoesNotClaimUnrelatedSameNamedDataModelTable()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(
            _fixture.BatchToken,
            sheetName,
            "A1:A2",
            [["Value"], [71]]));
        RequireSuccess(_tables.Create(_fixture.BatchToken, sheetName, "A", "A1:A2"));
        RequireSuccess(_tables.AddToDataModel(_fixture.BatchToken, "A"));
        _fixture.RegisterDataModelTableForCleanup("A");
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            "A",
            PrefixQueryMCode,
            PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup("A");

        var query = RequireSuccess(_queries.List(_fixture.BatchToken)).Queries.Single(item => item.Name == "A");
        Assert.True(query.IsConnectionOnly);
        Assert.Equal(
            PowerQueryLoadMode.ConnectionOnly,
            RequireSuccess(_queries.GetLoadConfig(_fixture.BatchToken, "A")).LoadMode);
        var error = Assert.Throws<InvalidOperationException>(
            () => _queries.Refresh(
                _fixture.BatchToken,
                "A",
                TimeSpan.FromSeconds(30)));
        Assert.Contains("Could not find connection or table", error.Message);
        Assert.Contains(
            RequireSuccess(_dataModel.ListTables(_fixture.BatchToken)).Tables,
            table => table.Name == "A");
        Assert.Equal(PrefixQueryMCode, RequireSuccess(_queries.View(_fixture.BatchToken, "A")).MCode);
        PowerQueryStateAssertions.AssertRows([[71]],
            RequireSuccess(_dataModel.Evaluate(_fixture.BatchToken, "EVALUATE 'A'")).Rows);
        PowerQueryStateAssertions.AssertRows([[71]], RequireSuccess(_commands.GetValues(
            _fixture.BatchToken, sheetName, "A2")).Values);
    }

    [Fact]
    public async Task Evaluate_Success_SaveAndReopen_PersistsNoTemporaryArtifacts()
    {
        CreatePrefixQueries(PowerQueryLoadMode.LoadToTable);
        RequireSuccess(_connections.Create(
            _fixture.BatchToken,
            "Connection",
            "ODBC;DSN=PreservedGenericConnection"));
        _fixture.RegisterConnectionForCleanup("Connection");

        var result = RequireSuccess(_queries.Evaluate(_fixture.BatchToken, PrefixQueryMCode));
        Assert.Equal(PrefixQueryMCode, result.MCode);
        Assert.Equal(["Value"], result.Columns);
        Assert.Equal(1, result.RowCount);
        Assert.Equal(1, result.ColumnCount);
        PowerQueryStateAssertions.AssertRows([[1]], result.Rows);
        AssertNoEvaluateArtifactsAfterReopen();

        await _fixture.SaveAndReopenAsync();

        AssertNoEvaluateArtifactsAfterReopen();
    }

    [Fact]
    public async Task Evaluate_Failure_SaveAndReopen_PersistsNoTemporaryArtifacts()
    {
        const string invalidMCode = "let Source = UndefinedFunction() in Source";
        CreatePrefixQueries(PowerQueryLoadMode.LoadToTable);
        RequireSuccess(_connections.Create(
            _fixture.BatchToken,
            "Connection",
            "ODBC;DSN=PreservedGenericConnection"));
        _fixture.RegisterConnectionForCleanup("Connection");

        var error = Assert.Throws<InvalidOperationException>(
            () => _queries.Evaluate(_fixture.BatchToken, invalidMCode));
        Assert.Contains("UndefinedFunction", error.Message);
        AssertNoEvaluateArtifactsAfterReopen();

        await _fixture.SaveAndReopenAsync();

        AssertNoEvaluateArtifactsAfterReopen();
    }

    [Fact]
    public void Evaluate_PreservesQueryConnectionAliasedByTemporaryDisplayName()
    {
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            "Connection",
            NeighborQueryMCode,
            PowerQueryLoadMode.LoadToTable,
            "QueryData"));
        _fixture.RegisterPowerQueryForCleanup("Connection");
        _fixture.RegisterSheetForCleanup("QueryData");

        var result = RequireSuccess(_queries.Evaluate(_fixture.BatchToken, PrefixQueryMCode));

        Assert.True(result.Success);
        Assert.Equal(
            PowerQueryLoadMode.LoadToTable,
            RequireSuccess(_queries.GetLoadConfig(_fixture.BatchToken, "Connection")).LoadMode);
        Assert.Equal(
            "QueryData",
            RequireSuccess(_queries.GetLoadConfig(_fixture.BatchToken, "Connection")).TargetSheet);
        PowerQueryStateAssertions.AssertRows([[1]], result.Rows);
        Assert.Empty(FindEvaluateArtifacts());
        PowerQueryStateAssertions.AssertStored(_fixture, "Connection", NeighborQueryMCode,
            PowerQueryLoadMode.LoadToTable, "QueryData", ["Value"], [[23], [47]]);
    }

    private void CreatePrefixQueries(PowerQueryLoadMode aaLoadMode)
    {
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            "A",
            PrefixQueryMCode,
            PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup("A");
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            "AA",
            NeighborQueryMCode,
            aaLoadMode,
            aaLoadMode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth
                ? "AAData"
                : null));
        _fixture.RegisterPowerQueryForCleanup("AA");
        if (aaLoadMode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth)
        {
            _fixture.RegisterSheetForCleanup("AAData");
        }
        AssertPrefixStored();
        PowerQueryStateAssertions.AssertStored(_fixture, "AA", NeighborQueryMCode,
            aaLoadMode, aaLoadMode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth
                ? "AAData" : null, ["Value"], [[23], [47]]);
    }

    private void AssertPrefixStored() =>
        PowerQueryStateAssertions.AssertStored(_fixture, "A", PrefixQueryMCode,
            PowerQueryLoadMode.ConnectionOnly, null, ["Value"], [[1]]);

    private void AssertWorksheetLoadPreserved(string queryName) =>
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, NeighborQueryMCode,
            PowerQueryLoadMode.LoadToTable, "AAData", ["Value"], [[23], [47]]);

    private void AssertNoEvaluateArtifactsAfterReopen()
    {
        Assert.Empty(FindEvaluateArtifacts());
        Assert.Contains(
            RequireSuccess(_connections.List(_fixture.BatchToken)).Connections,
            connection => connection.Name == "Connection");
        AssertPrefixStored();
        AssertWorksheetLoadPreserved("AA");
        AssertGenericConnection();
    }

    private List<string> FindEvaluateArtifacts() =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            var artifacts = new List<string>();
            Excel.Queries? queries = null;
            Excel.Sheets? worksheets = null;
            Excel.Connections? connections = null;

            try
            {
                queries = ctx.Book.Queries;
                for (var index = 1; index <= queries.Count; index++)
                {
                    Excel.WorkbookQuery? query = null;
                    try
                    {
                        query = queries.Item(index);
                        if (query.Name.StartsWith(
                                "__pq_eval_",
                                StringComparison.OrdinalIgnoreCase))
                        {
                            artifacts.Add($"query:{query.Name}");
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref query);
                    }
                }

                worksheets = ctx.Book.Worksheets;
                for (var index = 1; index <= worksheets.Count; index++)
                {
                    Excel.Worksheet? worksheet = null;
                    try
                    {
                        worksheet = (Excel.Worksheet)worksheets.Item[index];
                        if (worksheet.Name.StartsWith(
                                "__pq_eval_",
                                StringComparison.OrdinalIgnoreCase))
                        {
                            artifacts.Add($"worksheet:{worksheet.Name}");
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref worksheet);
                    }
                }

                connections = ctx.Book.Connections;
                for (var index = 1; index <= connections.Count; index++)
                {
                    Excel.WorkbookConnection? connection = null;
                    Excel.OLEDBConnection? oleDbConnection = null;
                    try
                    {
                        connection = connections.Item(index);
                        if (connection.Type != Excel.XlConnectionType.xlConnectionTypeOLEDB)
                        {
                            continue;
                        }

                        oleDbConnection = connection.OLEDBConnection;
                        var connectionString =
                            Convert.ToString(oleDbConnection.Connection)
                            ?? string.Empty;
                        if (connectionString.Contains(
                                "Location=__pq_eval_",
                                StringComparison.OrdinalIgnoreCase))
                        {
                            artifacts.Add($"connection:{connection.Name}");
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref oleDbConnection);
                        ComUtilities.Release(ref connection);
                    }
                }

                return artifacts;
            }
            finally
            {
                ComUtilities.Release(ref connections);
                ComUtilities.Release(ref worksheets);
                ComUtilities.Release(ref queries);
            }
        });
}
