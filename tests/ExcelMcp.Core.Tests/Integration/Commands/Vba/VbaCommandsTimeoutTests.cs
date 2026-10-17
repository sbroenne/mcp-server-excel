using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Vba;

[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Collection("Sequential")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "VBA")]
public sealed class VbaCommandsTimeoutTests :
    IClassFixture<VbaTestsFixture>
{
    private readonly VbaTestsFixture _fixture;

    public VbaCommandsTimeoutTests(VbaTestsFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact(Timeout = 60000)]
    [Trait("RunType", "OnDemand")]
    [Trait("Speed", "Slow")]
    public async Task ScriptCommands_Run_WhenMacroExceedsCallerTimeout_PreservesExecutionAndLateChanges()
    {
        await Task.Yield();

        var testFile = _fixture.CreateTestFile();
        const string vbaCode = """
            Sub WaitForTimeout()
                Application.Wait Now + TimeSerial(0, 0, 15)
                ThisWorkbook.Worksheets(1).Range("A1").Value = "Macro finished"
            End Sub
            """;

        var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromMinutes(2),
            testFile);
        var commands = new VbaCommands();

        var failure = await Record.ExceptionAsync(async () =>
        {
            await Task.Yield();
            _ = commands.Import(batch, "TimeoutModule", vbaCode);

            var runStopwatch = Stopwatch.StartNew();
            var timeoutException = Assert.Throws<TimeoutException>(() =>
                commands.Run(
                    batch,
                    "TimeoutModule.WaitForTimeout",
                    TimeSpan.FromSeconds(1)));
            Assert.Equal("Timeout", OperationFailureClassifier.Classify(timeoutException));
            runStopwatch.Stop();

            Assert.InRange(
                runStopwatch.Elapsed,
                TimeSpan.FromMilliseconds(500),
                TimeSpan.FromSeconds(10));
            Assert.True(batch.HasTimedOutOperation);

            var retryStopwatch = Stopwatch.StartNew();
            var retryException =
                Assert.Throws<TimeoutException>(() => commands.List(batch));
            retryStopwatch.Stop();

            Assert.True(
                retryStopwatch.Elapsed < TimeSpan.FromSeconds(1),
                $"Pending execution should reject new work immediately, but took {retryStopwatch.Elapsed}.");
            Assert.Contains(
                "previous wait expired",
                retryException.Message,
                StringComparison.OrdinalIgnoreCase);
            Assert.Throws<ExcelBusyException>(batch.Dispose);
            Assert.True(batch.IsExcelProcessAlive());
            await WaitUntilReadyAsync(batch);
            Assert.False(batch.HasTimedOutOperation);
            Assert.Equal("Macro finished", batch.Execute((context, _) =>
            {
                Microsoft.Office.Interop.Excel.Sheets? sheets = null;
                Microsoft.Office.Interop.Excel.Worksheet? sheet = null;
                Microsoft.Office.Interop.Excel.Range? cell = null;
                try
                {
                    sheets = context.Book.Worksheets;
                    sheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets[1];
                    cell = sheet.Range["A1"];
                    return cell.Value2 as string;
                }
                finally
                {
                    ComUtilities.Release(ref cell);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                }
            }));
        });
        var cleanupFailure = await Record.ExceptionAsync(async () =>
        {
            if (batch.IsExcelProcessAlive())
                await WaitUntilReadyAsync(batch);
            var disposeStopwatch = Stopwatch.StartNew();
            batch.Dispose();
            disposeStopwatch.Stop();

            Assert.True(
                disposeStopwatch.Elapsed < TimeSpan.FromSeconds(30),
                $"Timed-out VBA cleanup took {disposeStopwatch.Elapsed}.");
        });
        var failures = new[] { failure, cleanupFailure }.OfType<Exception>().ToArray();
        if (failures.Length > 0)
            throw new AggregateException("VBA timeout preservation or cleanup failed.", failures);
    }

    private static async Task WaitUntilReadyAsync(IExcelBatch batch)
    {
        var state = Assert.IsAssignableFrom<IExcelBatchRefreshState>(batch);
        var timer = Stopwatch.StartNew();
        while (state.GetRefreshState() != WorkbookRefreshState.Ready)
        {
            Assert.True(timer.Elapsed < TimeSpan.FromSeconds(30), "VBA execution did not finish.");
            await Task.Delay(100);
        }
    }
}
