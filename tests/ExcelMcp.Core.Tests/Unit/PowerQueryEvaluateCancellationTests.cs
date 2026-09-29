using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
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

        public TimeSpan OperationTimeout => TimeSpan.FromMilliseconds(10);

        public bool IsExcelVisible => false;

        public CancellationToken CapturedToken { get; private set; }

        public bool CancelDuringExecution { get; init; }

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
            if (CancelDuringExecution)
            {
                Assert.True(
                    SpinWait.SpinUntil(
                        () => cancellationToken.IsCancellationRequested,
                        TimeSpan.FromSeconds(1)),
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
}
