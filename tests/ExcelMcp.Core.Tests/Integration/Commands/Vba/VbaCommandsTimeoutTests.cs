using System.Diagnostics;
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
    public async Task ScriptCommands_Run_WhenMacroExceedsCallerTimeout_ReportsTimeoutAndPoisonsBatch()
    {
        await Task.Yield();

        var testFile = _fixture.CreateTestFile();
        const string vbaCode = """
            Sub WaitForTimeout()
                Application.Wait Now + TimeSerial(0, 0, 15)
            End Sub
            """;

        var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromMinutes(2),
            testFile);
        var commands = new VbaCommands();

        try
        {
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
                $"A poisoned batch should fail immediately, but took {retryStopwatch.Elapsed}.");
            Assert.Contains(
                "previous operation",
                retryException.Message,
                StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            var disposeStopwatch = Stopwatch.StartNew();
            batch.Dispose();
            disposeStopwatch.Stop();

            Assert.True(
                disposeStopwatch.Elapsed < TimeSpan.FromSeconds(30),
                $"Timed-out VBA cleanup took {disposeStopwatch.Elapsed}.");
        }
    }
}
