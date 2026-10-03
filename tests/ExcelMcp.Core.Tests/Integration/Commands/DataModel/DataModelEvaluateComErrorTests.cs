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
        AssertModelRows(innerBatch);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var batch = new InjectedCancellationBatch(
            innerBatch,
            cancellation.Token);
        var commands = new DataModelCommands();

        Assert.ThrowsAny<OperationCanceledException>(() =>
            commands.Evaluate(batch, "EVALUATE ROW(\"Probe\", 42)"));
        AssertModelRows(innerBatch);
    }

    [Fact]
    public void Evaluate_CancelledDuringLargeResult_StopsExtraction()
    {
        using var innerBatch = ExcelSession.BeginBatch(fixture.TestFilePath);
        AssertModelRows(innerBatch);
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
        AssertModelRows(innerBatch);
    }

    [Fact]
    public void Evaluate_InvalidDax_PreservesComExceptionTopology()
    {
        using var batch = ExcelSession.BeginBatch(fixture.TestFilePath);
        var commands = new DataModelCommands();
        AssertModelRows(batch);

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
        AssertModelRows(batch);
    }

    private static void AssertModelRows(IExcelBatch batch)
    {
        var result = new DataModelCommands().Evaluate(batch,
            """
            EVALUATE SELECTCOLUMNS('SalesTable',
                "SalesID", 'SalesTable'[SalesID], "Amount", 'SalesTable'[Amount])
            ORDER BY [SalesID]
            """);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal(["[SalesID]", "[Amount]"], result.Columns);
        Assert.Equal(10, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        decimal[] amounts = [150, 250, 175, 300, 125, 450, 200, 350, 275, 180];
        for (var row = 0; row < amounts.Length; row++)
        {
            Assert.Equal(row + 1, Convert.ToDecimal(result.Rows[row][0],
                System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(amounts[row], Convert.ToDecimal(result.Rows[row][1],
                System.Globalization.CultureInfo.InvariantCulture));
        }
    }
}
