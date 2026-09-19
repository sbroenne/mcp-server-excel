using System.Globalization;
using System.Runtime.InteropServices;
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

    private readonly IPowerQueryCommands _queries =
        fixture.CreateCommands<IPowerQueryCommands>();
    private readonly IConnectionCommands _connections =
        fixture.CreateCommands<IConnectionCommands>();
    private readonly ITableCommands _tables =
        fixture.CreateCommands<ITableCommands>();
    private readonly IDataModelCommands _dataModel =
        fixture.CreateCommands<IDataModelCommands>();

    [Fact]
    public async Task ExactIdentity_ReadAndRefreshPaths_DoNotTreatAAAsA()
    {
        CreatePrefixQueries(PowerQueryLoadMode.LoadToTable);

        var listResult = _queries.List(_fixture.BatchToken);
        Assert.True(listResult.Queries.Single(query => query.Name == "A").IsConnectionOnly);
        Assert.False(listResult.Queries.Single(query => query.Name == "AA").IsConnectionOnly);

        var viewResult = _queries.View(_fixture.BatchToken, "a");
        Assert.True(viewResult.IsConnectionOnly);

        var loadConfig = _queries.GetLoadConfig(_fixture.BatchToken, "a");
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, loadConfig.LoadMode);

        var response = await _fixture.SendForFailureAsync(
            "powerquery.refresh",
            new { queryName = "a", timeout = TimeSpan.FromSeconds(30) });
        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal("Prerequisite", response.ErrorCategory);
        Assert.Contains(
            "Could not find connection or table for query 'a'",
            response.ErrorMessage);

        AssertWorksheetLoadPreserved("AA");
    }

    [Fact]
    public void LoadTo_PrefixQuery_PreservesAAWorksheetDestination()
    {
        CreatePrefixQueries(PowerQueryLoadMode.LoadToTable);

        var result = _queries.LoadTo(
            _fixture.BatchToken,
            "A",
            PowerQueryLoadMode.LoadToTable,
            "AData",
            "A1");
        _fixture.RegisterSheetForCleanup("AData");

        Assert.True(result.Success);
        Assert.Equal(
            PowerQueryLoadMode.LoadToTable,
            _queries.GetLoadConfig(_fixture.BatchToken, "A").LoadMode);
        AssertWorksheetLoadPreserved("AA");
    }

    [Fact]
    public void Unload_PrefixQuery_PreservesAAWorksheetDestination()
    {
        CreatePrefixQueries(PowerQueryLoadMode.LoadToTable);

        var result = _queries.Unload(_fixture.BatchToken, "A");

        Assert.True(result.Success);
        AssertWorksheetLoadPreserved("AA");
    }

    [Fact]
    public void Delete_PrefixQuery_PreservesAAWorksheetDestination()
    {
        CreatePrefixQueries(PowerQueryLoadMode.LoadToTable);

        var result = _queries.Delete(_fixture.BatchToken, "a");
        _fixture.ForgetPowerQuery("A");

        Assert.True(result.Success);
        Assert.DoesNotContain(
            _queries.List(_fixture.BatchToken).Queries,
            query => query.Name == "A");
        AssertWorksheetLoadPreserved("AA");
    }

    [Fact]
    public void Unload_PrefixQuery_PreservesAADataModelDestination()
    {
        CreatePrefixQueries(PowerQueryLoadMode.LoadToDataModel);

        var result = _queries.Unload(_fixture.BatchToken, "A");

        Assert.True(result.Success);
        var tables = _dataModel.ListTables(_fixture.BatchToken);
        Assert.Contains(tables.Tables, table => table.Name == "AA");
        Assert.Equal(
            PowerQueryLoadMode.LoadToDataModel,
            _queries.GetLoadConfig(_fixture.BatchToken, "AA").LoadMode);
    }

    [Fact]
    public void ConnectionOnlyQuery_DoesNotClaimUnrelatedSameNamedDataModelTable()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _commands.SetValues(
            _fixture.BatchToken,
            sheetName,
            "A1:A2",
            [["Value"], [1]]);
        _tables.Create(_fixture.BatchToken, sheetName, "A", "A1:A2");
        _tables.AddToDataModel(_fixture.BatchToken, "A");
        _fixture.RegisterDataModelTableForCleanup("A");
        _queries.Create(
            _fixture.BatchToken,
            "A",
            PrefixQueryMCode,
            PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup("A");

        var query = _queries.List(_fixture.BatchToken).Queries.Single(item => item.Name == "A");
        Assert.True(query.IsConnectionOnly);
        Assert.Equal(
            PowerQueryLoadMode.ConnectionOnly,
            _queries.GetLoadConfig(_fixture.BatchToken, "A").LoadMode);
        Assert.Throws<InvalidOperationException>(
            () => _queries.Refresh(
                _fixture.BatchToken,
                "A",
                TimeSpan.FromSeconds(30)));
        Assert.Contains(
            _dataModel.ListTables(_fixture.BatchToken).Tables,
            table => table.Name == "A");
    }

    [Fact]
    public async Task Evaluate_Success_SaveAndReopen_PersistsNoTemporaryArtifacts()
    {
        _connections.Create(
            _fixture.BatchToken,
            "Connection",
            "ODBC;DSN=PreservedGenericConnection");
        _fixture.RegisterConnectionForCleanup("Connection");

        var result = _queries.Evaluate(_fixture.BatchToken, PrefixQueryMCode);
        Assert.True(result.Success);

        await _fixture.SaveAndReopenAsync();

        AssertNoEvaluateArtifactsAfterReopen();
    }

    [Fact]
    public async Task Evaluate_Failure_SaveAndReopen_PersistsNoTemporaryArtifacts()
    {
        const string invalidMCode = "let Source = UndefinedFunction() in Source";
        _connections.Create(
            _fixture.BatchToken,
            "Connection",
            "ODBC;DSN=PreservedGenericConnection");
        _fixture.RegisterConnectionForCleanup("Connection");

        Assert.ThrowsAny<Exception>(
            () => _queries.Evaluate(_fixture.BatchToken, invalidMCode));

        await _fixture.SaveAndReopenAsync();

        AssertNoEvaluateArtifactsAfterReopen();
    }

    [Fact]
    public void Evaluate_PreservesQueryConnectionAliasedByTemporaryDisplayName()
    {
        _queries.Create(
            _fixture.BatchToken,
            "Connection",
            PrefixQueryMCode,
            PowerQueryLoadMode.LoadToTable,
            "QueryData");
        _fixture.RegisterPowerQueryForCleanup("Connection");
        _fixture.RegisterSheetForCleanup("QueryData");

        var result = _queries.Evaluate(_fixture.BatchToken, PrefixQueryMCode);

        Assert.True(result.Success);
        Assert.Equal(
            PowerQueryLoadMode.LoadToTable,
            _queries.GetLoadConfig(_fixture.BatchToken, "Connection").LoadMode);
        Assert.Equal(
            "QueryData",
            _queries.GetLoadConfig(_fixture.BatchToken, "Connection").TargetSheet);
    }

    private void CreatePrefixQueries(PowerQueryLoadMode aaLoadMode)
    {
        _queries.Create(
            _fixture.BatchToken,
            "A",
            PrefixQueryMCode,
            PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup("A");
        _queries.Create(
            _fixture.BatchToken,
            "AA",
            PrefixQueryMCode,
            aaLoadMode,
            aaLoadMode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth
                ? "AAData"
                : null);
        _fixture.RegisterPowerQueryForCleanup("AA");
        if (aaLoadMode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth)
        {
            _fixture.RegisterSheetForCleanup("AAData");
        }
    }

    private void AssertWorksheetLoadPreserved(string queryName)
    {
        var config = _queries.GetLoadConfig(_fixture.BatchToken, queryName);
        Assert.Equal(PowerQueryLoadMode.LoadToTable, config.LoadMode);
        Assert.Equal("AAData", config.TargetSheet);

        var view = _queries.View(_fixture.BatchToken, queryName);
        Assert.False(view.IsConnectionOnly);
    }

    private void AssertNoEvaluateArtifactsAfterReopen()
    {
        Assert.Empty(FindEvaluateArtifacts());
        Assert.Contains(
            _connections.List(_fixture.BatchToken).Connections,
            connection => connection.Name == "Connection");
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
                        if (Convert.ToInt32(
                                connection.Type,
                                CultureInfo.InvariantCulture) != 1)
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
                    catch (COMException)
                    {
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
