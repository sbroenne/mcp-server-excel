using System.Diagnostics;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelBatch")]
[Trait("RunType", "OnDemand")]
[Trait("RequiresExcel", "true")]
[Collection("Sequential")]
public sealed class ExcelBatchTimeoutTests : IAsyncLifetime
{
    private readonly string _path = Path.Join(Path.GetTempPath(), $"batch-timeout-{Guid.NewGuid():N}.xlsx");

    public Task InitializeAsync()
    {
        var template = Path.Join(AppContext.BaseDirectory, "Integration", "Session", "TestFiles", "batch-test-static.xlsx");
        File.Copy(template, _path);
        return Task.CompletedTask;
    }

    public Task DisposeAsync()
    {
        File.Delete(_path);
        return Task.CompletedTask;
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Execute_WaitExpires_PreservesUnsavedWorkAndRecoversAfterCompletion(bool cancelled, bool lateFailure)
    {
        using var owned = new OwnedExcelProcessScope();
        using var release = new ManualResetEventSlim();
        using var cancellation = new CancellationTokenSource();
        var batch = Open();
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var originalPid = batch.ExcelProcessId;
        var failure = await Record.ExceptionAsync(async () =>
        {
            var execution = Task.Run(() => batch.Execute((ctx, token) =>
            {
                Assert.False(token.CanBeCanceled);
                WriteCell(ctx, "unsaved-checkpoint");
                started.SetResult();
                Assert.True(release.Wait(TimeSpan.FromSeconds(20)));
                if (lateFailure)
                    throw new InvalidOperationException("Late operation failure.");
                return 1;
            }, cancelled ? cancellation.Token : default));
            await started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            if (cancelled)
            {
                cancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => execution);
            }
            else
            {
                var error = await Assert.ThrowsAsync<TimeoutException>(() => execution);
                Assert.Contains("remains open", error.Message, StringComparison.Ordinal);
            }
            Assert.True(batch.IsExcelProcessAlive());
            Assert.Equal(originalPid, batch.ExcelProcessId);
            Assert.True(batch.HasTimedOutOperation);
            var pending = batch.OperationStatus;
            Assert.Equal("running", pending.State);
            Assert.True(pending.ExecutionPending);
            Assert.True(pending.WaitExpired);
            Assert.True(pending.ElapsedSeconds > 0);
            Assert.True(pending.WindowReveal is "shown" or "showRequested" or "alreadyVisible");

            var rejected = false;
            var timer = Stopwatch.StartNew();
            Assert.Throws<TimeoutException>(() => batch.Execute((_, _) => { rejected = true; }));
            Assert.False(rejected);
            Assert.True(timer.Elapsed < TimeSpan.FromSeconds(1));
            Assert.Throws<ExcelBusyException>(batch.Dispose);
            Assert.True(batch.IsExcelProcessAlive());

            release.Set();
            await WaitForReadyAsync(batch);
            var completed = batch.OperationStatus;
            Assert.Equal(lateFailure ? "failed" : "completed", completed.State);
            Assert.False(completed.ExecutionPending);
            Assert.True(completed.WaitExpired);
            Assert.Equal(lateFailure ? "Late operation failure." : null, completed.ErrorMessage);
            Assert.True(batch.IsExcelVisible);
            Assert.False(batch.HasTimedOutOperation);
            Assert.Equal("unsaved-checkpoint", batch.Execute((ctx, _) => ReadCell(ctx)));
        });
        release.Set();
        await FinishAsync(batch, owned, failure);
    }

    [Fact]
    public async Task Execute_QueuedWaitExpires_DoesNotExecuteOrInvalidateRunningWork()
    {
        using var owned = new OwnedExcelProcessScope();
        using var release = new ManualResetEventSlim();
        using var lifetime = new CancellationTokenSource();
        var batch = Open();
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var first = Task.Run(() => batch.Execute((_, _) =>
        {
            started.SetResult();
            Assert.True(release.Wait(TimeSpan.FromSeconds(20)));
            return 1;
        }, lifetime.Token));
        var failure = await Record.ExceptionAsync(async () =>
        {
            await started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            var ran = false;
            var error = Assert.Throws<TimeoutException>(() => batch.Execute((_, _) => { ran = true; }));
            Assert.Contains("session queue", error.Message, StringComparison.Ordinal);
            Assert.False(batch.HasTimedOutOperation);
            Assert.False(batch.OperationStatus.WaitExpired);
            Assert.Equal("running", batch.OperationStatus.State);
            release.Set();
            Assert.Equal(1, await first.WaitAsync(TimeSpan.FromSeconds(10)));
            Assert.Equal(3, batch.Execute((_, _) => 3));
            Assert.False(ran);
        });
        release.Set();
        var joinFailure = await Record.ExceptionAsync(async () => await first.WaitAsync(TimeSpan.FromSeconds(10)));
        await FinishAsync(batch, owned, failure, joinFailure);
    }

    [Fact]
    public async Task GetRefreshState_StartedScan_IsBoundedWithoutCancellingScan()
    {
        using var owned = new OwnedExcelProcessScope();
        using var release = new ManualResetEventSlim();
        var batch = Open();
        batch.BeforeRefreshStateReadHookForTests = () => Assert.True(release.Wait(TimeSpan.FromSeconds(10)));
        var failure = Record.Exception(() =>
        {
            var timer = Stopwatch.StartNew();
            Assert.Equal(WorkbookRefreshState.Unknown, batch.GetRefreshState());
            Assert.True(timer.Elapsed < TimeSpan.FromSeconds(1));
            Assert.False(batch.HasTimedOutOperation);
            Assert.Equal(WorkbookRefreshState.Unknown, batch.GetRefreshState());
            Assert.Throws<ExcelBusyException>(batch.Dispose);
        });
        batch.BeforeRefreshStateReadHookForTests = null;
        release.Set();
        await FinishAsync(batch, owned, failure);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Startup_ExpiredWait_RetainsSessionPathAndOriginalOpen(bool readOnlyValidation)
    {
        using var owned = new OwnedExcelProcessScope();
        using var release = new ManualResetEventSlim();
        using var manager = new SessionManager();
        ExcelBatch.BeforeWorkbookOpenHook = (_, _) => Assert.True(release.Wait(TimeSpan.FromSeconds(20)));
        ExcelBatch? batch = null;
        var failure = await Record.ExceptionAsync(async () =>
        {
            var error = Assert.Throws<SessionStartupPendingException>(() =>
            {
                if (readOnlyValidation)
                    manager.ValidateWorkbookOpen(_path, TimeSpan.FromSeconds(2));
                else
                    manager.CreateSessionWithTimeouts(_path, false, TimeSpan.FromSeconds(2), TimeSpan.FromSeconds(2));
            });
            batch = Assert.IsType<ExcelBatch>(manager.GetSession(error.SessionId));
            Assert.True(batch.IsExcelProcessAlive());
            Assert.Equal(error.SessionId, Assert.Single(manager.GetActiveSessions()).SessionId);
            Assert.False(manager.ValidateClose(error.SessionId).CanClose);
            var status = manager.GetOperationStatus(error.SessionId)!;
            Assert.Equal("starting", status.State);
            Assert.True(status.ExecutionPending);
            Assert.True(status.WaitExpired);
            Assert.Throws<InvalidOperationException>(() => manager.CreateSession(_path));
            release.Set();
            await WaitForReadyAsync(batch);
            Assert.Equal("completed", manager.GetOperationStatus(error.SessionId)!.State);
            Assert.True(manager.GetOperationStatus(error.SessionId)!.WaitExpired);
            Assert.Equal(_path, batch.Execute((ctx, _) => ctx.Book.FullName));
            Assert.Equal(readOnlyValidation, batch.Execute((ctx, _) => ctx.Book.ReadOnly));
            Assert.True(manager.CloseSession(error.SessionId));
            Assert.Empty(manager.GetActiveSessions());
        });
        ExcelBatch.BeforeWorkbookOpenHook = null;
        release.Set();
        if (batch is not null)
            await FinishAsync(batch, owned, failure);
        else if (failure is not null)
            throw failure;
    }

    [Fact]
    public async Task Execute_ExcelPromptTimesOut_ReportsPromptAndPreservesWorkbook()
    {
        using var owned = new OwnedExcelProcessScope();
        var batch = Assert.IsType<ExcelBatch>(ExcelSession.BeginBatchWithTimeouts(
            false, TimeSpan.FromSeconds(10), ComInteropConstants.DefaultOperationTimeout, _path));
        var identity = Assert.Single(SessionManager.GetTrackedExcelProcesses(),
            item => item.ProcessId == batch.ExcelProcessId);
        var failure = Record.Exception(() =>
        {
            var timeout = Assert.Throws<TimeoutException>(() =>
                batch.Execute((ctx, _) => ctx.App.InputBox(
                    "The test database is unavailable.", "SSAS connection test", Type: 2)));
            Assert.Contains("dialog open", timeout.Message, StringComparison.Ordinal);
            Assert.Contains("SSAS connection test", timeout.Message, StringComparison.Ordinal);
            Assert.Contains("Dialog text", timeout.Message, StringComparison.Ordinal);
            Assert.Contains("does not itself cancel", timeout.Message, StringComparison.Ordinal);
            Assert.True(batch.IsExcelProcessAlive());
            Assert.True(batch.HasTimedOutOperation);
        });
        var dismissFailure = Record.Exception(() =>
        {
            Assert.True(OwnedProcessGuard.IsAlive(identity));
            var posted = false;
            EnumWindows((window, _) =>
            {
                if (GetWindowThreadProcessId(window, out var pid) == 0)
                    return true;
                var title = new char[256];
                var length = GetWindowText(window, title, title.Length);
                if (pid == identity.ProcessId && new string(title, 0, length) == "SSAS connection test")
                {
                    posted = PostMessage(window, 0x0010, nint.Zero, nint.Zero);
                    return false;
                }
                return true;
            }, nint.Zero);
            Assert.True(posted, "The test-owned prompt could not be dismissed.");
        });
        await FinishAsync(batch, owned, failure, dismissFailure);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BeginBatch_DefaultOpenTimeout_Is120Seconds(bool show)
    {
        using var batch = ExcelSession.BeginBatch(show, operationTimeout: null, _path);
        Assert.Equal(TimeSpan.FromSeconds(120), batch.OperationTimeout);
        Assert.NotNull(batch.ExcelProcessId);
        Assert.True(batch.IsExcelProcessAlive());
        Assert.Equal(show, batch.IsExcelVisible);
        Assert.Equal(_path, batch.Execute((context, _) => context.Book.FullName), ignoreCase: true);
    }

    [Fact]
    public void BeginBatch_OpenTimeoutOverride_IsHonored()
    {
        using var batch = ExcelSession.BeginBatch(false, TimeSpan.FromSeconds(45), _path);
        Assert.Equal(TimeSpan.FromSeconds(45), batch.OperationTimeout);
        Assert.Equal(_path, batch.Execute((context, _) => context.Book.FullName), ignoreCase: true);
    }

    [Fact]
    public void BeginBatch_SeparateStartupTimeout_AllowsStartupLongerThanOperationTimeout()
    {
        ExcelBatch.BeforeWorkbookOpenHook = (_, _) => Thread.Sleep(TimeSpan.FromSeconds(2));
        try
        {
            using var batch = ExcelSession.BeginBatchWithTimeouts(false, TimeSpan.FromSeconds(1),
                ComInteropConstants.DefaultOperationTimeout, _path);
            Assert.Equal(TimeSpan.FromSeconds(1), batch.OperationTimeout);
            Assert.Equal(_path, batch.Execute((context, _) => context.Book.FullName), ignoreCase: true);
        }
        finally
        {
            ExcelBatch.BeforeWorkbookOpenHook = null;
        }
    }

    private ExcelBatch Open() => Assert.IsType<ExcelBatch>(ExcelSession.BeginBatchWithTimeouts(
        false, TimeSpan.FromSeconds(2), ComInteropConstants.DefaultOperationTimeout, _path));

    private static async Task WaitForReadyAsync(ExcelBatch batch)
    {
        var timer = Stopwatch.StartNew();
        while (batch.GetRefreshState() != WorkbookRefreshState.Ready)
        {
            Assert.True(timer.Elapsed < TimeSpan.FromSeconds(15), "Excel did not become ready after work returned.");
            await Task.Delay(100);
        }
    }

    private static async Task FinishAsync(ExcelBatch batch, OwnedExcelProcessScope owned, params Exception?[] failures)
    {
        var cleanupFailure = await Record.ExceptionAsync(async () =>
        {
            if (batch.IsExcelProcessAlive())
                await WaitForReadyAsync(batch);
            batch.Dispose();
        });
        var processFailure = Record.Exception(() => owned.AssertAllExited());
        var all = failures.Append(cleanupFailure).Append(processFailure).OfType<Exception>().ToArray();
        if (all.Length > 0)
            throw new AggregateException("Timeout preservation or cleanup failed.", all);
    }

    private static void WriteCell(ExcelContext context, string value)
    {
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? cell = null;
        try
        {
            sheets = context.Book.Worksheets;
            sheet = (Excel.Worksheet)sheets[1];
            cell = sheet.Range["A1"];
            cell.Value2 = value;
        }
        finally
        {
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
        }
    }

    private static string ReadCell(ExcelContext context)
    {
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? cell = null;
        try
        {
            sheets = context.Book.Worksheets;
            sheet = (Excel.Worksheet)sheets[1];
            cell = sheet.Range["A1"];
            return Assert.IsType<string>((object?)cell.Value2);
        }
        finally
        {
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
        }
    }

    private delegate bool WindowCallback(nint window, nint parameter);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(WindowCallback callback, nint parameter);
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint window, out uint processId);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(nint window, [Out] char[] text, int count);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PostMessage(nint window, uint message, nint wParam, nint lParam);
}
