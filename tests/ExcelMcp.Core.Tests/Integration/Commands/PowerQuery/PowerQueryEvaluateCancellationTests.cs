using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.PowerQuery;

[Collection("PowerQuery")]
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
[Trait("Speed", "Slow")]
public sealed class PowerQueryEvaluateCancellationTests(
    PowerQueryTestsFixture fixture)
{
    [Fact]
    public Task Evaluate_CallerCancellation_CleansUpTemporaryObjectsAndPreservesExistingData() =>
        VerifyCancellationAsync(duringMExecution: false);

    [Fact]
    public Task Evaluate_CancellationDuringMExecution_CleansUpTemporaryObjectsAndPreservesExistingData() =>
        VerifyCancellationAsync(duringMExecution: true);

    private async Task VerifyCancellationAsync(bool duringMExecution)
    {
        var mCode = """
            let
                Source = Function.InvokeAfter(
                    () => #table({"Value"}, {{1}}),
                    #duration(0, 0, 0, 2))
            in
                Source
            """;
        var testFile = fixture.CreateTestFile();
        using var innerBatch = ExcelSession.BeginBatch(testFile);
        var commands = new PowerQueryCommands(new DataModelCommands());
        var seeded = commands.Create(innerBatch, "RetainedQuery",
            "#table(type table [Value = number], {{17}, {29}})", PowerQueryLoadMode.LoadToTable);
        Assert.True(seeded.Success, seeded.ErrorMessage);
        var initialState = GetObjectState(innerBatch);
        using var cancellation = new CancellationTokenSource();
        var source = duringMExecution ? new LocalMSourceCancellationProbe(cancellation) : null;
        var failures = new List<Exception>();
        try
        {
            if (source is not null)
                mCode = $"Table.PromoteHeaders(Csv.Document(Binary.Buffer(Web.Contents(\"{source.Url}\", [Timeout = #duration(0, 0, 0, 10)])), [Delimiter = \",\", Columns = 1, Encoding = 65001]))";
            else
                cancellation.CancelAfter(TimeSpan.FromMilliseconds(250));
            using var batch = new InjectedCancellationBatch(
                innerBatch,
                cancellation.Token);
            Assert.ThrowsAny<OperationCanceledException>(() => commands.Evaluate(batch, mCode));
            if (source is not null)
            {
                Assert.True(source.Requests > 0, "Excel must request the M source before cancellation.");
                Assert.True(source.CancelledBeforeResponse, "Cancellation must precede the M-source response.");
            }

            Assert.True(cancellation.IsCancellationRequested);
            var afterCancellation = GetObjectState(innerBatch);
            Assert.Equal(initialState, afterCancellation);

            var followUp = commands.Evaluate(
                innerBatch,
                "let Source = #table({\"Value\"}, {{42}}) in Source");
            Assert.True(followUp.Success, followUp.ErrorMessage);
            Assert.Null(followUp.ErrorMessage);
            Assert.Equal("Value", Assert.Single(followUp.Columns));
            Assert.Equal(1, followUp.RowCount);
            Assert.Equal(1, followUp.ColumnCount);
            Assert.Equal(42d, Assert.Single(Assert.Single(followUp.Rows)));
            Assert.Equal(afterCancellation, GetObjectState(innerBatch));
        }
        catch (Exception ex)
        {
            failures.Add(ex);
        }
        finally
        {
            if (source is not null)
            {
                try { await source.DisposeAsync(); }
                catch (Exception cleanupFailure) { failures.Add(cleanupFailure); }
            }
        }
        if (failures.Count != 0)
            throw new AggregateException("Cancellation verification or local source cleanup failed.", failures);
    }

    private static string GetObjectState(
        IExcelBatch batch) =>
        batch.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Queries? queries = null;
            Excel.Connections? connections = null;
            Excel.Worksheet? guard = null;
            Excel.Range? cells = null;
            try
            {
                sheets = context.Book.Worksheets;
                queries = context.Book.Queries;
                connections = context.Book.Connections;
                var sheetNames = new List<string>();
                for (var index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet? sheet = null;
                    try
                    {
                        sheet = (Excel.Worksheet)sheets[index];
                        sheetNames.Add(sheet.Name);
                    }
                    finally { ComUtilities.Release(ref sheet); }
                }
                var definitions = new Dictionary<string, string>();
                for (var index = 1; index <= queries.Count; index++)
                {
                    Excel.WorkbookQuery? query = null;
                    try
                    {
                        query = queries.Item(index);
                        definitions.Add(query.Name, query.Formula);
                    }
                    finally { ComUtilities.Release(ref query); }
                }
                var connectionIdentities = new Dictionary<string, string>();
                for (var index = 1; index <= connections.Count; index++)
                {
                    Excel.WorkbookConnection? connection = null;
                    try
                    {
                        connection = connections.Item(index);
                        Assert.True(Sbroenne.ExcelMcp.Core.PowerQuery.PowerQueryHelpers.TryGetMashupLocation(
                            connection, out var location));
                        connectionIdentities.Add(connection.Name, location);
                    }
                    finally { ComUtilities.Release(ref connection); }
                }
                guard = (Excel.Worksheet)sheets["RetainedQuery"];
                cells = guard.Range["A1:A3"];
                var values = Assert.IsType<object[,]>((object?)cells.Value2);
                Assert.Equal("Value", values[1, 1]);
                Assert.Equal(17d, values[2, 1]);
                Assert.Equal(29d, values[3, 1]);
                return JsonSerializer.Serialize(new
                {
                    Sheets = sheetNames,
                    Queries = definitions,
                    Connections = connectionIdentities,
                    Values = values.Cast<object?>().ToArray()
                });
            }
            finally
            {
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref guard);
                ComUtilities.Release(ref connections);
                ComUtilities.Release(ref queries);
                ComUtilities.Release(ref sheets);
            }
        });
}
