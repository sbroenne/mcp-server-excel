using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

[Collection("DataModel")]
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Slow")]
public sealed class DataModelEvaluateComErrorTests(
    DataModelPivotTableFixture fixture)
{
    [Fact]
    public void Evaluate_CancelledBeforeExecution_StopsBeforeQuery()
    {
        using var innerBatch = ExcelSession.BeginBatch(fixture.TestFilePath);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var batch = new InjectedCancellationBatch(
            innerBatch,
            cancellation.Token);
        var commands = new DataModelCommands();

        Assert.ThrowsAny<OperationCanceledException>(() =>
            commands.Evaluate(batch, "EVALUATE ROW(\"Probe\", 42)"));
    }

    [Fact]
    public void Evaluate_CancelledDuringLargeResult_StopsExtraction()
    {
        using var innerBatch = ExcelSession.BeginBatch(fixture.TestFilePath);
        using var cancellation = new CancellationTokenSource();
        using var batch = new InjectedCancellationBatch(
            innerBatch,
            cancellation.Token);
        bool extractionStarted = false;
        var commands = new DataModelCommands(() =>
        {
            extractionStarted = true;
            cancellation.Cancel();
        });

        Assert.ThrowsAny<OperationCanceledException>(() =>
            commands.Evaluate(
                batch,
                """
                EVALUATE
                CROSSJOIN(
                    SELECTCOLUMNS('SalesTable', "A", 'SalesTable'[SalesID]),
                    SELECTCOLUMNS('SalesTable', "B", 'SalesTable'[SalesID]),
                    SELECTCOLUMNS('SalesTable', "C", 'SalesTable'[SalesID]),
                    SELECTCOLUMNS('SalesTable', "D", 'SalesTable'[SalesID]),
                    SELECTCOLUMNS('SalesTable', "E", 'SalesTable'[SalesID])
                )
                """));
        Assert.True(extractionStarted);
        Assert.True(cancellation.IsCancellationRequested);

        var followUp = commands.Evaluate(
            innerBatch,
            "EVALUATE ROW(\"Probe\", 42)");
        Assert.True(followUp.Success, followUp.ErrorMessage);
        Assert.Equal(42m, Assert.Single(Assert.Single(followUp.Rows)));
    }

    [Fact]
    public void Evaluate_InvalidDax_PreservesComExceptionTopology()
    {
        using var batch = ExcelSession.BeginBatch(fixture.TestFilePath);
        var commands = new DataModelCommands();

        var exception = Assert.ThrowsAny<Exception>(() =>
            commands.Evaluate(batch, "EVALUATE INVALID_FUNCTION()"));

        Assert.Equal(
            "ComInterop",
            OperationFailureClassifier.Classify(exception));
        Assert.Contains(
            "DAX evaluation failed",
            exception.Message,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "INVALID_FUNCTION",
            exception.Message,
            StringComparison.Ordinal);
        var cause =
            Assert.IsType<System.Runtime.InteropServices.COMException>(
                exception.InnerException);
        Assert.Contains(
            cause.Message,
            exception.Message,
            StringComparison.Ordinal);
    }
}
