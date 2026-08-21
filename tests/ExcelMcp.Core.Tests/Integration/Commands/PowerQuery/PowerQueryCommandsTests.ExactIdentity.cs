using System.Runtime.InteropServices;
using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

public partial class PowerQueryCommandsTests
{
    private const string PrefixQueryMCode = "let Source = #table({\"Value\"}, {{1}}) in Source";

    [Fact]
    public void ExactIdentity_ReadAndRefreshPaths_DoNotTreatAAAsA()
    {
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);
        CreatePrefixQueries(batch, PowerQueryLoadMode.LoadToTable);

        var listResult = _powerQueryCommands.List(batch);
        Assert.True(listResult.Queries.Single(query => query.Name == "A").IsConnectionOnly);
        Assert.False(listResult.Queries.Single(query => query.Name == "AA").IsConnectionOnly);

        var viewResult = _powerQueryCommands.View(batch, "a");
        Assert.True(viewResult.IsConnectionOnly);

        var loadConfig = _powerQueryCommands.GetLoadConfig(batch, "a");
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, loadConfig.LoadMode);

        var exception = Assert.Throws<InvalidOperationException>(
            () => _powerQueryCommands.Refresh(batch, "a", TimeSpan.FromSeconds(30)));
        Assert.Contains("Could not find connection or table for query 'a'", exception.Message);

        AssertWorksheetLoadPreserved(batch, "AA");
    }

    [Fact]
    public void LoadTo_PrefixQuery_PreservesAAWorksheetDestination()
    {
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);
        CreatePrefixQueries(batch, PowerQueryLoadMode.LoadToTable);

        var result = _powerQueryCommands.LoadTo(
            batch,
            "A",
            PowerQueryLoadMode.LoadToTable,
            "AData",
            "A1");

        Assert.True(result.Success);
        Assert.Equal(PowerQueryLoadMode.LoadToTable, _powerQueryCommands.GetLoadConfig(batch, "A").LoadMode);
        AssertWorksheetLoadPreserved(batch, "AA");
    }

    [Fact]
    public void Unload_PrefixQuery_PreservesAAWorksheetDestination()
    {
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);
        CreatePrefixQueries(batch, PowerQueryLoadMode.LoadToTable);

        var result = _powerQueryCommands.Unload(batch, "A");

        Assert.True(result.Success);
        AssertWorksheetLoadPreserved(batch, "AA");
    }

    [Fact]
    public void Delete_PrefixQuery_PreservesAAWorksheetDestination()
    {
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);
        CreatePrefixQueries(batch, PowerQueryLoadMode.LoadToTable);

        var result = _powerQueryCommands.Delete(batch, "a");

        Assert.True(result.Success);
        Assert.DoesNotContain(_powerQueryCommands.List(batch).Queries, query => query.Name == "A");
        AssertWorksheetLoadPreserved(batch, "AA");
    }

    [Fact]
    public void Unload_PrefixQuery_PreservesAADataModelDestination()
    {
        var testFile = _fixture.CreateTestFile();
        var dataModelCommands = new DataModelCommands();

        using var batch = ExcelSession.BeginBatch(testFile);
        CreatePrefixQueries(batch, PowerQueryLoadMode.LoadToDataModel);

        var result = _powerQueryCommands.Unload(batch, "A");

        Assert.True(result.Success);
        var tables = dataModelCommands.ListTables(batch);
        Assert.Contains(tables.Tables, table => table.Name == "AA");
        Assert.Equal(
            PowerQueryLoadMode.LoadToDataModel,
            _powerQueryCommands.GetLoadConfig(batch, "AA").LoadMode);
    }

    [Fact]
    public void ConnectionOnlyQuery_DoesNotClaimUnrelatedSameNamedDataModelTable()
    {
        var testFile = _fixture.CreateTestFile();
        var rangeCommands = new RangeCommands();
        var tableCommands = new TableCommands();
        var dataModelCommands = new DataModelCommands();

        using var batch = ExcelSession.BeginBatch(testFile);
        rangeCommands.SetValues(batch, "Sheet1", "A1:A2", [["Value"], [1]]);
        tableCommands.Create(batch, "Sheet1", "A", "A1:A2");
        tableCommands.AddToDataModel(batch, "A");
        _powerQueryCommands.Create(batch, "A", PrefixQueryMCode, PowerQueryLoadMode.ConnectionOnly);

        var query = _powerQueryCommands.List(batch).Queries.Single(item => item.Name == "A");
        Assert.True(query.IsConnectionOnly);
        Assert.Equal(
            PowerQueryLoadMode.ConnectionOnly,
            _powerQueryCommands.GetLoadConfig(batch, "A").LoadMode);
        Assert.Throws<InvalidOperationException>(
            () => _powerQueryCommands.Refresh(batch, "A", TimeSpan.FromSeconds(30)));
        Assert.Contains(dataModelCommands.ListTables(batch).Tables, table => table.Name == "A");
    }

    [Fact]
    public void Evaluate_Success_SaveAndReopen_PersistsNoTemporaryArtifacts()
    {
        var testFile = _fixture.CreateTestFile();
        var connectionCommands = new ConnectionCommands();

        using (var batch = ExcelSession.BeginBatch(testFile))
        {
            connectionCommands.Create(batch, "Connection", "ODBC;DSN=PreservedGenericConnection");
            var result = _powerQueryCommands.Evaluate(batch, PrefixQueryMCode);
            Assert.True(result.Success);
            batch.Save();
        }

        AssertNoEvaluateArtifactsAfterReopen(testFile, connectionCommands);
    }

    [Fact]
    public void Evaluate_Failure_SaveAndReopen_PersistsNoTemporaryArtifacts()
    {
        var testFile = _fixture.CreateTestFile();
        var connectionCommands = new ConnectionCommands();
        const string invalidMCode = "let Source = UndefinedFunction() in Source";

        using (var batch = ExcelSession.BeginBatch(testFile))
        {
            connectionCommands.Create(batch, "Connection", "ODBC;DSN=PreservedGenericConnection");
            Assert.ThrowsAny<Exception>(() => _powerQueryCommands.Evaluate(batch, invalidMCode));
            batch.Save();
        }

        AssertNoEvaluateArtifactsAfterReopen(testFile, connectionCommands);
    }

    [Fact]
    public void Evaluate_PreservesQueryConnectionAliasedByTemporaryDisplayName()
    {
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);
        _powerQueryCommands.Create(
            batch,
            "Connection",
            PrefixQueryMCode,
            PowerQueryLoadMode.LoadToTable,
            "QueryData");

        var result = _powerQueryCommands.Evaluate(batch, PrefixQueryMCode);

        Assert.True(result.Success);
        Assert.Equal(
            PowerQueryLoadMode.LoadToTable,
            _powerQueryCommands.GetLoadConfig(batch, "Connection").LoadMode);
        Assert.Equal(
            "QueryData",
            _powerQueryCommands.GetLoadConfig(batch, "Connection").TargetSheet);
    }

    private void CreatePrefixQueries(IExcelBatch batch, PowerQueryLoadMode aaLoadMode)
    {
        _powerQueryCommands.Create(batch, "A", PrefixQueryMCode, PowerQueryLoadMode.ConnectionOnly);
        _powerQueryCommands.Create(
            batch,
            "AA",
            PrefixQueryMCode,
            aaLoadMode,
            aaLoadMode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth
                ? "AAData"
                : null);
    }

    private void AssertWorksheetLoadPreserved(IExcelBatch batch, string queryName)
    {
        var config = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.Equal(PowerQueryLoadMode.LoadToTable, config.LoadMode);
        Assert.Equal("AAData", config.TargetSheet);

        var view = _powerQueryCommands.View(batch, queryName);
        Assert.False(view.IsConnectionOnly);
    }

    private static void AssertNoEvaluateArtifactsAfterReopen(
        string testFile,
        ConnectionCommands connectionCommands)
    {
        using var reopenedBatch = ExcelSession.BeginBatch(testFile);
        var artifacts = FindEvaluateArtifacts(reopenedBatch);

        Assert.Empty(artifacts);
        Assert.Contains(
            connectionCommands.List(reopenedBatch).Connections,
            connection => connection.Name == "Connection");
    }

    private static List<string> FindEvaluateArtifacts(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            var artifacts = new List<string>();
            Excel.Queries? queries = null;
            Excel.Sheets? worksheets = null;
            Excel.Connections? connections = null;

            try
            {
                queries = ctx.Book.Queries;
                for (int i = 1; i <= queries.Count; i++)
                {
                    Excel.WorkbookQuery? query = null;
                    try
                    {
                        query = queries.Item(i);
                        if (query.Name.StartsWith("__pq_eval_", StringComparison.OrdinalIgnoreCase))
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
                for (int i = 1; i <= worksheets.Count; i++)
                {
                    Excel.Worksheet? worksheet = null;
                    try
                    {
                        worksheet = (Excel.Worksheet)worksheets.Item[i];
                        if (worksheet.Name.StartsWith("__pq_eval_", StringComparison.OrdinalIgnoreCase))
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
                for (int i = 1; i <= connections.Count; i++)
                {
                    Excel.WorkbookConnection? connection = null;
                    Excel.OLEDBConnection? oleDbConnection = null;
                    try
                    {
                        connection = connections.Item(i);
                        if (Convert.ToInt32(connection.Type, CultureInfo.InvariantCulture) != 1)
                        {
                            continue;
                        }

                        oleDbConnection = connection.OLEDBConnection;
                        var connectionString = Convert.ToString(oleDbConnection.Connection) ?? string.Empty;
                        if (connectionString.Contains(
                            "Location=__pq_eval_",
                            StringComparison.OrdinalIgnoreCase))
                        {
                            artifacts.Add($"connection:{connection.Name}");
                        }
                    }
                    catch (COMException)
                    {
                        // Non-OLEDB connection subtype.
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
}
