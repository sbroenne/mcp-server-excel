using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
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
    public void Evaluate_CancelledDuringRefresh_CleansUpTemporaryObjects()
    {
        const string delayedMCode = """
            let
                Source = Function.InvokeAfter(
                    () => #table({"Value"}, {{1}}),
                    #duration(0, 0, 0, 2))
            in
                Source
            """;
        var testFile = fixture.CreateTestFile();
        using var innerBatch = ExcelSession.BeginBatch(testFile);
        var initialState = GetObjectCounts(innerBatch);
        using var cancellation = new CancellationTokenSource();
        cancellation.CancelAfter(TimeSpan.FromMilliseconds(250));
        using var batch = new InjectedCancellationBatch(
            innerBatch,
            cancellation.Token);
        var commands = new PowerQueryCommands(new DataModelCommands());

        Assert.ThrowsAny<OperationCanceledException>(() =>
            commands.Evaluate(batch, delayedMCode));
        Assert.True(cancellation.IsCancellationRequested);
        Assert.Equal(initialState, GetObjectCounts(innerBatch));
    }

    private static (int Sheets, int Queries, int Connections) GetObjectCounts(
        IExcelBatch batch) =>
        batch.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Queries? queries = null;
            Excel.Connections? connections = null;
            try
            {
                sheets = context.Book.Worksheets;
                queries = context.Book.Queries;
                connections = context.Book.Connections;
                return (sheets.Count, queries.Count, connections.Count);
            }
            finally
            {
                ComUtilities.Release(ref connections);
                ComUtilities.Release(ref queries);
                ComUtilities.Release(ref sheets);
            }
        });
}
