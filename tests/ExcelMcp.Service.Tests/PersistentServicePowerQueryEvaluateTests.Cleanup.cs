using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePowerQueryEvaluateTests
{
    [Fact]
    public void Evaluate_InvalidMCode_ThrowsError()
    {
        const string code = "let Source = UndefinedFunction() in Source";
        var guard = CreateEvaluationGuard();
        var before = SnapshotNativeArtifacts();
        var error = Assert.Throws<InvalidOperationException>(() =>
            _queries.Evaluate(_fixture.BatchToken, code));
        Assert.Contains("Expression.Error", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("powerquery.evaluate failed [Expression/PowerQueryCommandException]", error.Message);
        Assert.Contains("UndefinedFunction", error.Message);
        Assert.Equal(before, SnapshotNativeArtifacts());
        AssertEvaluationGuard(guard);
        EvaluateChecked("let Source = #table({\"Recovered\"}, {{71}}) in Source", ["Recovered"], [[71]]);
        AssertEvaluationGuard(guard);
    }

    [Fact]
    public void Evaluate_AfterExecution_CleansUpTempObjects() =>
        EvaluateChecked("let Source = #table({\"X\"}, {{1}}) in Source", ["X"], [[1]]);

    private string SnapshotNativeArtifacts() =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            var artifacts = new List<string[]>();
            Excel.Sheets? sheets = null;
            Excel.Queries? queries = null;
            Excel.Connections? connections = null;
            try
            {
                sheets = context.Book.Worksheets;
                for (var index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet? sheet = null;
                    Excel.ListObjects? tables = null;
                    Excel.QueryTables? queryTables = null;
                    try
                    {
                        sheet = (Excel.Worksheet)sheets.Item[index];
                        artifacts.Add(["sheet", sheet.Name]);
                        tables = sheet.ListObjects;
                        for (var tableIndex = 1; tableIndex <= tables.Count; tableIndex++)
                        {
                            Excel.ListObject? table = null;
                            try
                            {
                                table = tables.Item[tableIndex];
                                artifacts.Add(["table", sheet.Name, table.Name]);
                            }
                            finally { ComUtilities.Release(ref table); }
                        }
                        queryTables = sheet.QueryTables;
                        for (var queryIndex = 1; queryIndex <= queryTables.Count; queryIndex++)
                        {
                            Excel.QueryTable? queryTable = null;
                            Excel.WorkbookConnection? connection = null;
                            try
                            {
                                queryTable = queryTables.Item(queryIndex);
                                connection = queryTable.WorkbookConnection;
                                artifacts.Add(["querytable", sheet.Name, queryTable.Name, connection.Name]);
                                Assert.False(queryTable.Refreshing);
                            }
                            finally
                            {
                                ComUtilities.Release(ref connection);
                                ComUtilities.Release(ref queryTable);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref queryTables);
                        ComUtilities.Release(ref tables);
                        ComUtilities.Release(ref sheet);
                    }
                }
                queries = context.Book.Queries;
                for (var index = 1; index <= queries.Count; index++)
                {
                    Excel.WorkbookQuery? query = null;
                    try
                    {
                        query = queries.Item(index);
                        artifacts.Add(["query", query.Name, query.Formula]);
                    }
                    finally { ComUtilities.Release(ref query); }
                }
                connections = context.Book.Connections;
                for (var index = 1; index <= connections.Count; index++)
                {
                    Excel.WorkbookConnection? connection = null;
                    Excel.OLEDBConnection? oledb = null;
                    try
                    {
                        connection = connections.Item(index);
                        var source = "";
                        if (connection.Type == Excel.XlConnectionType.xlConnectionTypeOLEDB)
                        {
                            oledb = connection.OLEDBConnection;
                            source = Convert.ToString(oledb.Connection, CultureInfo.InvariantCulture) ?? "";
                            Assert.False(oledb.Refreshing);
                        }
                        artifacts.Add(["connection", connection.Name, connection.Type.ToString(),
                            connection.InModel.ToString(), source]);
                    }
                    finally
                    {
                        ComUtilities.Release(ref oledb);
                        ComUtilities.Release(ref connection);
                    }
                }
                return JsonSerializer.Serialize(artifacts.OrderBy(
                    artifact => JsonSerializer.Serialize(artifact), StringComparer.Ordinal));
            }
            finally
            {
                ComUtilities.Release(ref connections);
                ComUtilities.Release(ref queries);
                ComUtilities.Release(ref sheets);
            }
        });
}
