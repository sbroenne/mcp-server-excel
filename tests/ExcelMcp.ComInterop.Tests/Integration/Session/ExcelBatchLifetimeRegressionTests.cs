using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Channels;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelBatch")]
[Trait("RequiresExcel", "true")]
[Trait("RunType", "OnDemand")]
[Collection("Sequential")]
public sealed class ExcelBatchLifetimeRegressionTests : IDisposable
{
    private readonly string _directory = Path.GetFullPath(Path.Combine("TestResults", $"batch-lifetime-{Guid.NewGuid():N}"));
    private readonly string _path;

    public ExcelBatchLifetimeRegressionTests()
    {
        Directory.CreateDirectory(_directory);
        _path = Path.Combine(_directory, "workbook.xlsx");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Integration", "Session", "TestFiles", "batch-test-static.xlsx"), _path);
    }

    [Fact]
    public async Task Execute_QueuedOperationTimesOut_DoesNotRunCallback()
    {
        using var batch = ExcelSession.BeginBatch(false, TimeSpan.FromSeconds(8), _path);
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        using var callerTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(40));
        var first = Task.Run(() => batch.Execute((_, _) =>
        {
            entered.Set();
            Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
        }, callerTimeout.Token));
        var executed = 0;
        try
        {
            Assert.True(entered.Wait(TimeSpan.FromSeconds(10)));
            Assert.Throws<TimeoutException>(() => batch.Execute((_, _) => Interlocked.Increment(ref executed)));
        }
        finally
        {
            release.Set();
            await first.WaitAsync(TimeSpan.FromSeconds(10));
            // Disposal joins the STA, so the assertion covers every queued callback.
            batch.Dispose();
        }
        Assert.Equal(0, Volatile.Read(ref executed));
    }

    [Fact]
    public async Task Execute_PreviousOperationTimesOut_RejectsAlreadyQueuedCallback()
    {
        using var batch = ExcelSession.BeginBatch(false, TimeSpan.FromSeconds(8), _path);
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        using var callerTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(40));
        var first = Task.Run(() => Record.Exception(() => batch.Execute((_, _) =>
        {
            entered.Set();
            Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
        })));
        Task<Exception>? queued = null;
        var executed = 0;
        try
        {
            Assert.True(entered.Wait(TimeSpan.FromSeconds(10)));
            queued = Task.Run(() => Record.Exception(() =>
                batch.Execute((_, _) => Interlocked.Increment(ref executed), callerTimeout.Token)));
            Assert.True(SpinWait.SpinUntil(() => WorkQueue(batch).Reader.TryPeek(out _), TimeSpan.FromSeconds(5)));
            Assert.IsType<TimeoutException>(await first.WaitAsync(TimeSpan.FromSeconds(15)));
        }
        finally
        {
            release.Set();
        }
        Assert.IsType<TimeoutException>(await queued!.WaitAsync(TimeSpan.FromSeconds(10)));
        Assert.Equal(0, Volatile.Read(ref executed));
    }

    [Fact]
    public async Task Execute_DisposedWithQueuedOperation_DoesNotRunCallback()
    {
        using var batch = ExcelSession.BeginBatch(_path);
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var first = Task.Run(() => Record.Exception(() => batch.Execute((_, _) =>
        {
            entered.Set();
            Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
        })));
        Task<Exception>? queued = null;
        Task? disposal = null;
        var executed = 0;
        try
        {
            Assert.True(entered.Wait(TimeSpan.FromSeconds(10)));
            queued = Task.Run(() => Record.Exception(() =>
                batch.Execute((_, _) => Interlocked.Increment(ref executed))));
            Assert.True(SpinWait.SpinUntil(() => WorkQueue(batch).Reader.TryPeek(out _), TimeSpan.FromSeconds(10)));
            disposal = Task.Run(batch.Dispose);
            var disposed = typeof(ExcelBatch).GetField("_disposed", BindingFlags.Instance | BindingFlags.NonPublic)!;
            Assert.True(SpinWait.SpinUntil(() => (int)disposed.GetValue(batch)! != 0, TimeSpan.FromSeconds(10)));
        }
        finally
        {
            release.Set();
            await first.WaitAsync(TimeSpan.FromSeconds(10));
            if (disposal != null) await disposal.WaitAsync(TimeSpan.FromSeconds(60));
        }
        Assert.NotNull(await queued!.WaitAsync(TimeSpan.FromSeconds(10)));
        Assert.Equal(0, Volatile.Read(ref executed));
    }

    [Fact]
    public void Execute_SessionTimeout_CancelsCallbackToken()
    {
        using var batch = ExcelSession.BeginBatch(false, TimeSpan.FromSeconds(8), _path);
        using var observedCancellation = new ManualResetEventSlim();
        Assert.Throws<TimeoutException>(() => batch.Execute((_, token) =>
        {
            if (token.WaitHandle.WaitOne(TimeSpan.FromSeconds(20)))
                observedCancellation.Set();
            token.ThrowIfCancellationRequested();
        }));
        Assert.True(observedCancellation.Wait(TimeSpan.FromSeconds(5)), "The callback did not receive the session timeout.");
    }

    [Fact]
    public async Task Execute_TimedOutWaiter_CallbackCanStillUseCancellationToken()
    {
        using var batch = ExcelSession.BeginBatch(false, TimeSpan.FromSeconds(8), _path);
        using var release = new ManualResetEventSlim();
        var observed = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        try
        {
            Assert.Throws<TimeoutException>(() => batch.Execute((_, token) =>
            {
                try
                {
                    Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
                    observed.TrySetResult(token.WaitHandle.WaitOne(TimeSpan.Zero));
                }
                catch (Exception ex)
                {
                    observed.TrySetException(ex);
                }
            }));
        }
        finally
        {
            release.Set();
        }
        Assert.True(await observed.Task.WaitAsync(TimeSpan.FromSeconds(10)));
    }

    [Fact]
    public void Execute_OperationThrowsCancellation_DoesNotReportSessionTimeout()
    {
        using var batch = ExcelSession.BeginBatch(_path);
        using var operationCancellation = new CancellationTokenSource();
        operationCancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => batch.Execute((_, _) =>
            operationCancellation.Token.ThrowIfCancellationRequested()));
    }

    [Fact]
    public void Execute_CallerTimeoutLongerThanSessionTimeout_IsNotShortened()
    {
        using var batch = ExcelSession.BeginBatch(false, TimeSpan.FromSeconds(8), _path);
        using var callerTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        Assert.Equal(42, batch.Execute((_, token) =>
        {
            Assert.False(token.WaitHandle.WaitOne(TimeSpan.FromSeconds(9)));
            return 42;
        }, callerTimeout.Token));
        Assert.False(batch.HasTimedOutOperation);
    }

    [Fact]
    public void BeginBatch_GenericExcelError_PreservesOriginalFailure()
    {
        var expected = Assert.IsType<COMException>(Marshal.GetExceptionForHR(unchecked((int)0x800A03EC)));
        ExcelBatch.BeforeWorkbookOpenHook = (_, _) => throw expected;
        try
        {
            Assert.Same(expected, Assert.Throws<COMException>(() => ExcelSession.BeginBatch(_path)));
        }
        finally
        {
            ExcelBatch.BeforeWorkbookOpenHook = null;
        }
    }

    [Fact]
    public void CreateNew_SaveAsFails_PreservesOriginalComFailure()
    {
        var directoryPath = Path.Combine(_directory, "directory.xlsx");
        Directory.CreateDirectory(directoryPath);
        Assert.Throws<COMException>(() => ExcelSession.CreateNew(directoryPath, false, (_, _) => 0));
        // Failure cleanup must release the directory and any partially created workbook.
        Directory.Delete(directoryPath);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CreateNew_InitialFileSaved_CallbackChangesRequireExplicitSave(bool save)
    {
        var path = Path.Combine(_directory, "new-directory", "created.xlsx");
        Assert.Equal(42, ExcelSession.CreateNew(path, false, (ctx, _) =>
        {
            Assert.True(File.Exists(path));
            ctx.Book.Title = "Changed by callback";
            if (save) ctx.Book.Save();
            return 42;
        }));
        using var reopened = ExcelSession.BeginBatch(path);
        Assert.Equal(save ? "Changed by callback" : string.Empty,
            reopened.Execute((ctx, _) => Convert.ToString(ctx.Book.Title) ?? string.Empty));
    }

    public void Dispose() => Directory.Delete(_directory, recursive: true);

    private static Channel<Func<Task>> WorkQueue(IExcelBatch batch) =>
        (Channel<Func<Task>>)typeof(ExcelBatch)
            .GetField("_workQueue", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(batch)!;
}
