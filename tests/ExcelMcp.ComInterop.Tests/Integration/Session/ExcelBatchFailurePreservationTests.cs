using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

[Collection("Sequential")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelBatch")]
[Trait("RequiresExcel", "true")]
[Trait("RunType", "OnDemand")]
[System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201",
    Justification = "Synthetic getter errors exercise the real ExcelBatch STA work-item and caller boundary without depending on an Excel crash.")]
public sealed class ExcelBatchFailurePreservationTests : IDisposable
{
    private readonly string _directory = Path.Join(
        Path.GetTempPath(), $"execution-preservation-{Guid.NewGuid():N}");
    private readonly string _file;

    public ExcelBatchFailurePreservationTests()
    {
        Directory.CreateDirectory(_directory);
        _file = Path.Combine(_directory, "workbook.xlsx");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Integration", "Session", "TestFiles",
            "batch-test-static.xlsx"), _file);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Execute_OperationAndVisibilityFail_PreservesBothThroughCaller(
        bool canceledOperation, bool cancellableCaller)
    {
        RunWithBatch(batch =>
        {
            using var operationCancellation = new CancellationTokenSource();
            operationCancellation.Cancel();
            using var callerLifetime = new CancellationTokenSource();
            var cause = new InvalidOperationException("Original operation cause");
            Exception original = canceledOperation
                ? new OperationCanceledException("Operation canceled", cause, operationCancellation.Token)
                : cause;
            var snapshotFailure = new COMException("Visible getter failed", unchecked((int)0x80070005));
            var callerToken = cancellableCaller ? callerLifetime.Token : CancellationToken.None;

            var received = Record.Exception(() => WithFailedVisibility(batch, snapshotFailure,
                () => batch.Execute<int>((_, _) => throw original, callerToken)));

            AggregateException combined;
            if (canceledOperation)
            {
                var canceled = Assert.IsType<OperationCanceledException>(received);
                Assert.Equal(operationCancellation.Token, canceled.CancellationToken);
                Assert.Equal(original.Message, canceled.Message);
                combined = Assert.IsType<AggregateException>(canceled.InnerException);
                Assert.Same(cause, original.InnerException);
            }
            else
            {
                combined = Assert.IsType<AggregateException>(received);
            }

            Assert.Collection(combined.InnerExceptions,
                failure => Assert.Same(original, failure),
                failure => Assert.Same(snapshotFailure, failure));
            Assert.False(batch.HasTimedOutOperation);
            Assert.Equal(42, batch.Execute((_, _) => 42, callerLifetime.Token));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Execute_OperationCancellation_PreservesOriginalWithoutInventingTimeout(bool cancellableCaller)
    {
        RunWithBatch(batch =>
        {
            using var operationCancellation = new CancellationTokenSource();
            operationCancellation.Cancel();
            using var callerLifetime = new CancellationTokenSource();
            var cause = new InvalidOperationException("Original cancellation cause");
            var original = new OperationCanceledException("Operation canceled", cause, operationCancellation.Token);
            var callerToken = cancellableCaller ? callerLifetime.Token : CancellationToken.None;

            var received = Assert.Throws<OperationCanceledException>(() =>
                batch.Execute<int>((_, _) => throw original, callerToken));

            Assert.Same(original, received);
            Assert.Same(cause, received.InnerException);
            Assert.Equal(operationCancellation.Token, received.CancellationToken);
            Assert.False(batch.HasTimedOutOperation);
            Assert.Equal(42, batch.Execute((_, _) => 42, callerLifetime.Token));
        });
    }

    [Fact]
    public void Execute_OnlyVisibilityFails_PropagatesOriginalGetterError()
    {
        RunWithBatch(batch =>
        {
            var snapshotFailure = new COMException("Visible getter failed", unchecked((int)0x80070005));
            var received = Assert.Throws<COMException>(() =>
                WithFailedVisibility(batch, snapshotFailure, () => batch.Execute((_, _) => 42)));

            Assert.Same(snapshotFailure, received);
            Assert.False(batch.HasTimedOutOperation);
            Assert.Equal(42, batch.Execute((_, _) => 42));
        });
    }

    private static void WithFailedVisibility(ExcelBatch batch, Exception failure, Action operation)
    {
        var excelField = typeof(ExcelBatch).GetField("_excel", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(excelField);
        var originalExcel = excelField.GetValue(batch);
        Assert.NotNull(originalExcel);
        var failingExcel = DispatchProxy.Create<Excel.Application, FailingVisibilityProxy>();
        var proxy = (FailingVisibilityProxy)(object)failingExcel;
        proxy.Application = (Excel.Application)originalExcel;
        proxy.Failure = failure;
        Assert.Same(failure, Assert.Throws<COMException>(() => _ = failingExcel.Visible));
        excelField.SetValue(batch, failingExcel);
        try
        {
            operation();
        }
        finally
        {
            excelField.SetValue(batch, originalExcel);
            Assert.Equal(2, proxy.Reads);
        }
    }

    private void RunWithBatch(Action<ExcelBatch> test)
    {
        using var owned = new OwnedExcelProcessScope();
        var batch = Assert.IsType<ExcelBatch>(ExcelSession.BeginBatch(_file));
        var failures = new List<Exception>();
        try
        {
            test(batch);
        }
        catch (Exception failure)
        {
            failures.Add(failure);
        }
        finally
        {
            try { batch.Dispose(); }
            catch (Exception failure) { failures.Add(failure); }
            try { owned.AssertAllExited(); }
            catch (Exception failure) { failures.Add(failure); }
        }

        if (failures.Count == 1) ExceptionDispatchInfo.Capture(failures[0]).Throw();
        if (failures.Count > 1) throw new AggregateException("Execution regression or cleanup failed.", failures);
    }

    public void Dispose() => Directory.Delete(_directory, recursive: true);

    public class FailingVisibilityProxy : DispatchProxy
    {
        private int _reads;
        public Excel.Application? Application { get; set; }
        public Exception? Failure { get; set; }
        public int Reads => Volatile.Read(ref _reads);

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            switch (targetMethod?.Name)
            {
                case "get_Visible":
                    Interlocked.Increment(ref _reads);
                    throw Failure ?? new InvalidOperationException("The visibility failure was not configured.");
                case "get_ScreenUpdating":
                    return Application!.ScreenUpdating;
                case "get_Workbooks":
                    return Application!.Workbooks;
                case "set_ScreenUpdating":
                    Application!.ScreenUpdating = (bool)args![0]!;
                    return null;
                default:
                    throw new NotSupportedException(targetMethod?.Name);
            }
        }
    }
}
