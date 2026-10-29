using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PowerQuery")]
public sealed class PowerQueryEvaluateCancellationTests
{
    [Fact]
    public void Evaluate_UsesCancelableOperationToken()
    {
        using var batch = new CapturingBatch();
        var commands = new PowerQueryCommands(new DataModelCommands());

        Assert.Throws<ProbeException>(() =>
            commands.Evaluate(batch, "let Source = #table({\"Value\"}, {{1}}) in Source"));

        Assert.True(batch.CapturedToken.CanBeCanceled);
    }

    [Fact]
    public void Evaluate_WhenOperationTokenExpires_ReportsTimeout()
    {
        using var batch = new CapturingBatch
        {
            CancelDuringExecution = true
        };
        var commands = new PowerQueryCommands(new DataModelCommands());

        var exception = Assert.Throws<TimeoutException>(() =>
            commands.Evaluate(batch, "let Source = #table({\"Value\"}, {{1}}) in Source"));

        var cause = Assert.IsType<OperationCanceledException>(exception.InnerException);
        Assert.Equal(batch.CapturedToken, cause.CancellationToken);
        Assert.Contains("Power Query evaluate timed out", exception.Message);
    }

    private sealed class CapturingBatch : IExcelBatch
    {
        public string WorkbookPath => "unused.xlsx";

        public ILogger Logger => NullLogger.Instance;

        public IReadOnlyDictionary<string, Excel.Workbook> Workbooks { get; } =
            new Dictionary<string, Excel.Workbook>();

        public bool HasTimedOutOperation => false;

        public int? ExcelProcessId => null;

        public TimeSpan OperationTimeout => TimeSpan.FromMilliseconds(50);

        public bool IsExcelVisible => false;

        public CancellationToken CapturedToken { get; private set; }

        public bool CancelDuringExecution { get; init; }

        public Exception? ExecutionFailure { get; init; }

        public void Dispose()
        {
        }

        public void Execute(
            Action<ExcelContext, CancellationToken> operation,
            CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public T Execute<T>(
            Func<ExcelContext, CancellationToken, T> operation,
            CancellationToken cancellationToken = default)
        {
            CapturedToken = cancellationToken;
            if (ExecutionFailure is not null)
                throw ExecutionFailure;
            if (CancelDuringExecution)
            {
                Assert.True(
                    cancellationToken.WaitHandle.WaitOne(TimeSpan.FromSeconds(30)),
                    "The evaluation operation token did not expire.");
                throw new OperationCanceledException(cancellationToken);
            }
            throw new ProbeException();
        }

        public Excel.Workbook GetWorkbook(string filePath) =>
            throw new NotSupportedException();

        public bool IsExcelProcessAlive() => false;

        public void Save(CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public void UpdateWorkbookPath(string workbookPath) =>
            throw new NotSupportedException();
    }

    private sealed class ProbeException : Exception;

    [Fact]
    public void DataOperationWait_DeadlineExpires_ReportsTimeoutWithOriginalCause()
    {
        using var batch = new CapturingBatch { CancelDuringExecution = true };
        var failure = Assert.Throws<TimeoutException>(() =>
            ExcelBatchWait.Execute(batch, batch.OperationTimeout, (_, _) => 1));
        var cause = Assert.IsType<OperationCanceledException>(failure.InnerException);
        Assert.Equal(batch.CapturedToken, cause.CancellationToken);
        Assert.Equal("Timeout", OperationFailureClassifier.Classify(failure));
        Assert.Contains("wait expired", failure.Message, StringComparison.Ordinal);
        Assert.Contains("execution may still be pending", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DataOperationWait_ExcelCancels_PreservesExecutionFailure()
    {
        var cause = new OperationCanceledException("Excel cancelled execution.");
        using var batch = new CapturingBatch { ExecutionFailure = cause };
        Assert.Same(cause, Assert.Throws<OperationCanceledException>(() =>
            ExcelBatchWait.Execute(batch, TimeSpan.FromSeconds(5), (_, _) => 1)));
    }
}
