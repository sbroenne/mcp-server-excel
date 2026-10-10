using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Tests for the operation timeout → force-kill → cleanup chain in ExcelBatch.
///
/// These tests validate the fix for Bug 8 (Feb 2026) where a stuck IDispatch.Invoke
/// caused the MCP server to hang permanently because:
/// 1. No timeout recovery existed — ExcelBatch.Dispose() waited forever on STA thread join
/// 2. No pre-emptive kill — Excel process was never killed when operations timed out
/// 3. No session cleanup — WithSessionAsync didn't handle TimeoutException
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test that Execute() throws TimeoutException when operation exceeds timeout
/// - ✅ Test that _operationTimedOut triggers pre-emptive Process.Kill() in Dispose()
/// - ✅ Test that Dispose() completes (doesn't hang) after timeout
/// - ✅ Test that Excel process is cleaned up after timeout + dispose
/// - ✅ Test that cancelled operations also trigger cleanup
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelBatch")]
[Trait("RunType", "OnDemand")]
[Collection("Sequential")]
[Trait("RequiresExcel", "true")]
public class ExcelBatchTimeoutTests : IAsyncLifetime
{
    private readonly ITestOutputHelper _output;
    private static string? _staticTestFile;
    private string? _testFileCopy;

    public ExcelBatchTimeoutTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public Task InitializeAsync()
    {
        if (_staticTestFile == null)
        {
            var testFolder = Path.Join(AppContext.BaseDirectory, "Integration", "Session", "TestFiles");
            _staticTestFile = Path.Join(testFolder, "batch-test-static.xlsx");

            if (!File.Exists(_staticTestFile))
            {
                throw new FileNotFoundException($"Static test file not found at {_staticTestFile}.");
            }
        }

        _testFileCopy = Path.Join(Path.GetTempPath(), $"batch-timeout-test-{Guid.NewGuid():N}.xlsx");
        File.Copy(_staticTestFile, _testFileCopy, overwrite: true);

        return Task.Delay(500);
    }

    public Task DisposeAsync()
    {
        if (_testFileCopy != null && File.Exists(_testFileCopy))
        {
            File.Delete(_testFileCopy);
        }
        return Task.CompletedTask;
    }

    [Fact]
    public void BeginBatch_DefaultOpenTimeout_Is120Seconds()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        Assert.Equal(TimeSpan.FromSeconds(120), batch.OperationTimeout);
    }

    [Fact]
    public void GetRefreshState_StartedInspectionExceedsDeadline_RemainsUnknownWithoutPoisoningBatch()
    {
        using var owned = new OwnedExcelProcessScope();
        using var batch = ExcelSession.BeginBatchWithTimeouts(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(2),
            startupTimeout: ComInteropConstants.DefaultOperationTimeout,
            _testFileCopy!);
        var implementation = Assert.IsType<ExcelBatch>(batch);
        implementation.BeforeRefreshStateReadHookForTests = () => Thread.Sleep(TimeSpan.FromSeconds(4));
        var failure = Record.Exception(() =>
        {
            var elapsed = Stopwatch.StartNew();
            Assert.Equal(WorkbookRefreshState.Unknown, implementation.GetRefreshState());
            Assert.InRange(elapsed.Elapsed, TimeSpan.FromSeconds(2), TimeSpan.FromSeconds(6));
            Assert.False(batch.HasTimedOutOperation);
        });
        implementation.BeforeRefreshStateReadHookForTests = null;
        using var recoveryLifetime = new CancellationTokenSource();
        var recoveryFailure = Record.Exception(() =>
        {
            Assert.Equal(1, batch.Execute((_, _) => 1, recoveryLifetime.Token));
            Assert.Equal(WorkbookRefreshState.Ready, implementation.GetRefreshState());
        });
        var disposalFailure = Record.Exception(batch.Dispose);
        var processFailure = Record.Exception(() => owned.AssertAllExited());
        var failures = new[] { failure, recoveryFailure, disposalFailure, processFailure }
            .OfType<Exception>().ToArray();
        if (failures.Length > 0)
            throw new AggregateException("Readiness deadline regression or cleanup failed.", failures);
    }

    [Fact]
    public void BeginBatch_OpenTimeoutOverride_IsHonored()
    {
        using var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(45),
            _testFileCopy!);

        Assert.Equal(TimeSpan.FromSeconds(45), batch.OperationTimeout);
    }

    [Fact]
    public void BeginBatch_SeparateStartupTimeout_AllowsStartupLongerThanOperationTimeout()
    {
        ExcelBatch.BeforeWorkbookOpenHook = (_, _) => Thread.Sleep(TimeSpan.FromSeconds(2));

        try
        {
            using var batch = ExcelSession.BeginBatchWithTimeouts(
                show: false,
                operationTimeout: TimeSpan.FromSeconds(1),
                startupTimeout: ComInteropConstants.DefaultOperationTimeout,
                _testFileCopy!);

            Assert.Equal(TimeSpan.FromSeconds(1), batch.OperationTimeout);
        }
        finally
        {
            ExcelBatch.BeforeWorkbookOpenHook = null;
        }
    }

    [Fact]
    public void BeginBatch_StartupOpenBlocks_ThrowsTimeoutExceptionInsteadOfHanging()
    {
        using var startupBlocked = new ManualResetEventSlim(false);
        ExcelBatch.BeforeWorkbookOpenHook = (_, cancellationToken) =>
        {
            startupBlocked.Set();
            cancellationToken.WaitHandle.WaitOne(TimeSpan.FromSeconds(30));
        };

        try
        {
            var sw = Stopwatch.StartNew();

            var ex = Assert.Throws<TimeoutException>(() =>
                ExcelSession.BeginBatch(
                    show: false,
                    operationTimeout: TimeSpan.FromSeconds(8),
                    _testFileCopy!));

            sw.Stop();

            Assert.True(startupBlocked.Wait(TimeSpan.FromSeconds(10)),
                "Startup hook was not reached before the timeout assertion.");
            Assert.Contains("startup timed out", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(Path.GetFileName(_testFileCopy!), ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("timeout_seconds", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("show=true", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.True(sw.Elapsed < TimeSpan.FromSeconds(30),
                $"Startup timeout regression: BeginBatch took {sw.Elapsed.TotalSeconds:F1}s. Expected a bounded timeout, not a hang.");
        }
        finally
        {
            ExcelBatch.BeforeWorkbookOpenHook = null;
        }
    }

    [Fact]
    public void BeginBatch_IrmStartupOpenBlocks_TimeoutMessageIncludesInteractiveGuidance()
    {
        string fakeIrmFile = Path.Join(Path.GetTempPath(), $"batch-timeout-irm-{Guid.NewGuid():N}.xlsx");
        OleDataSpaceTestFile.Write(fakeIrmFile, "\tDRMDataSpace");

        ExcelBatch.BeforeWorkbookOpenHook = (_, cancellationToken) =>
        {
            cancellationToken.WaitHandle.WaitOne(TimeSpan.FromSeconds(30));
        };

        try
        {
            var ex = Assert.Throws<TimeoutException>(() =>
                ExcelSession.BeginBatch(
                    show: true,
                    operationTimeout: TimeSpan.FromSeconds(8),
                    fakeIrmFile));

            Assert.Contains("IRM/AIP-protected", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("show=true", ex.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            ExcelBatch.BeforeWorkbookOpenHook = null;

            File.Delete(fakeIrmFile);
        }
    }

    [Fact]
    public async Task Execute_QueuedOperationExpiresWithoutExecutingOrPoisoningBatch()
    {
        using var owned = new OwnedExcelProcessScope();
        using var releaseFirst = new ManualResetEventSlim();
        using var firstStarted = new ManualResetEventSlim();
        using var callerLifetime = new CancellationTokenSource();
        using var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(15),
            _testFileCopy!);

        var first = Task.Run(() => batch.Execute((_, _) =>
        {
            firstStarted.Set();
            Assert.True(releaseFirst.Wait(TimeSpan.FromSeconds(30)));
            return 1;
        }, callerLifetime.Token));
        var primaryFailure = await Record.ExceptionAsync(async () =>
        {
            Assert.True(firstStarted.Wait(TimeSpan.FromSeconds(5)));
            var expiredCallbackRan = false;
            var timeout = Assert.Throws<TimeoutException>(() => batch.Execute((_, _) =>
            {
                expiredCallbackRan = true;
                return 2;
            }));

            Assert.Contains("session queue", timeout.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(batch.HasTimedOutOperation);
            releaseFirst.Set();
            Assert.Equal(1, await first.WaitAsync(TimeSpan.FromSeconds(5)));
            Assert.False(expiredCallbackRan);
            Assert.Equal(3, batch.Execute((_, _) => 3));
        });
        releaseFirst.Set();
        var joinFailure = await Record.ExceptionAsync(async () => await first.WaitAsync(TimeSpan.FromSeconds(5)));
        var disposalFailure = Record.Exception(batch.Dispose);
        var processFailure = Record.Exception(() => owned.AssertAllExited());
        var failures = new[] { primaryFailure, joinFailure, disposalFailure, processFailure }
            .OfType<Exception>().ToArray();
        if (failures.Length > 0)
        {
            throw new AggregateException("Queued-operation test and cleanup failed.", failures);
        }
    }

    [Fact]
    public async Task Dispose_DiscardsQueuedOperationWithoutExecutingCallback()
    {
        using var owned = new OwnedExcelProcessScope();
        using var releaseFirst = new ManualResetEventSlim();
        using var firstStarted = new ManualResetEventSlim();
        using var secondQueued = new ManualResetEventSlim();
        using var callerLifetime = new CancellationTokenSource();
        using var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(30),
            _testFileCopy!);

        var first = Task.Run(() => batch.Execute((_, _) =>
        {
            firstStarted.Set();
            Assert.True(releaseFirst.Wait(TimeSpan.FromSeconds(10)));
            return 1;
        }, callerLifetime.Token));
        Task<int>? queued = null;
        Task? release = null;
        var primaryFailure = await Record.ExceptionAsync(async () =>
        {
            Assert.True(firstStarted.Wait(TimeSpan.FromSeconds(5)));
            ExcelBatch.WorkItemQueuedHookForTests = secondQueued.Set;
            var queuedCallbackRan = false;
            queued = Task.Run(() => batch.Execute((_, _) =>
            {
                queuedCallbackRan = true;
                return 2;
            }));
            Assert.True(secondQueued.Wait(TimeSpan.FromSeconds(5)));

            release = Task.Run(async () =>
            {
                await Task.Delay(100);
                releaseFirst.Set();
            });

            batch.Dispose();
            Assert.Equal(1, await first.WaitAsync(TimeSpan.FromSeconds(5)));
            await release.WaitAsync(TimeSpan.FromSeconds(5));
            await Assert.ThrowsAsync<ObjectDisposedException>(
                async () => await queued.WaitAsync(TimeSpan.FromSeconds(5)));
            Assert.False(queuedCallbackRan);
        });
        ExcelBatch.WorkItemQueuedHookForTests = null;
        releaseFirst.Set();
        var disposalFailure = Record.Exception(batch.Dispose);
        var joinFailure = await Record.ExceptionAsync(async () =>
        {
            await first.WaitAsync(TimeSpan.FromSeconds(5));
            if (release is not null)
            {
                await release.WaitAsync(TimeSpan.FromSeconds(5));
            }
            if (queued is not null)
            {
                await Assert.ThrowsAsync<ObjectDisposedException>(
                    async () => await queued.WaitAsync(TimeSpan.FromSeconds(5)));
            }
        });
        var processFailure = Record.Exception(() => owned.AssertAllExited());
        var failures = new[] { primaryFailure, disposalFailure, joinFailure, processFailure }
            .OfType<Exception>().ToArray();
        if (failures.Length > 0)
        {
            throw new AggregateException("Queued-disposal test and cleanup failed.", failures);
        }
    }

    /// <summary>
    /// REGRESSION TEST: Execute() must throw TimeoutException when operation exceeds the configured timeout.
    /// Before Bug 8 fix, timeout existed but had no recovery — the caller got the exception but
    /// Dispose() would then hang forever waiting for the STA thread.
    /// </summary>
    [Fact]
    public void Execute_OperationExceedsTimeout_ThrowsTimeoutException()
    {
        // Arrange — use a very short timeout (3 seconds) to trigger timeout quickly
        using var owned = new OwnedExcelProcessScope();
        using var batch = ExcelSession.BeginBatchWithTimeouts(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(3),
            startupTimeout: ComInteropConstants.DefaultOperationTimeout,
            _testFileCopy!);

        // Warm up — ensure Excel is ready
        Assert.Equal(_testFileCopy, batch.Execute((ctx, _) => ctx.Book.FullName));

        _output.WriteLine("Excel initialized, starting long-running operation...");

        // Act & Assert — operation that exceeds timeout must throw TimeoutException
        var sw = Stopwatch.StartNew();
        var ex = Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                // Simulate a hung operation — sleep longer than the timeout
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });
        sw.Stop();

        _output.WriteLine($"TimeoutException thrown after {sw.Elapsed.TotalSeconds:F1}s: {ex.Message}");
        Assert.Contains("timed out", ex.Message, StringComparison.OrdinalIgnoreCase);

        // Dispose must complete and not hang — this is the key regression test
        var disposeSw = Stopwatch.StartNew();
        batch.Dispose();
        disposeSw.Stop();

        _output.WriteLine($"Dispose completed in {disposeSw.Elapsed.TotalSeconds:F1}s");

        // Dispose should complete within a reasonable time (pre-emptive kill + join + wait)
        // Before the fix, Dispose() would hang forever here
        Assert.True(disposeSw.Elapsed < TimeSpan.FromSeconds(30),
            $"REGRESSION: Dispose() took {disposeSw.Elapsed.TotalSeconds:F1}s after timeout — " +
            "pre-emptive kill may not be working. Expected < 30s.");
        owned.AssertAllExited();
    }

    /// <summary>
    /// REGRESSION TEST: After timeout, the Excel process must be killed and cleaned up.
    /// Before Bug 8 fix, the hung Excel process would remain alive permanently.
    /// </summary>
    [Fact]
    public void Execute_AfterTimeout_ExcelProcessIsCleaned()
    {
        // Arrange
        using var owned = new OwnedExcelProcessScope();

        using var batch = ExcelSession.BeginBatchWithTimeouts(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(3),
            startupTimeout: ComInteropConstants.DefaultOperationTimeout,
            _testFileCopy!);

        // Get the Excel process ID before timeout
        int? excelPid = batch.ExcelProcessId;
        Assert.NotNull(excelPid);
        _output.WriteLine($"Excel PID for this session: {excelPid}");

        // Warm up
        Assert.Equal(_testFileCopy, batch.Execute((ctx, _) => ctx.Book.FullName));

        // Act — trigger timeout
        Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });

        // Dispose triggers pre-emptive kill
        batch.Dispose();

        owned.AssertAllExited();
    }

    /// <summary>
    /// REGRESSION TEST: Dispose after timeout must use shorter join timeout (aggressive cleanup).
    /// Before Bug 8 fix, Dispose() used the same 45-second join timeout even when the operation
    /// had already timed out, causing unnecessary delays.
    /// </summary>
    [Fact]
    public void Dispose_AfterTimeout_CompletesWithinAggressiveTimeout()
    {
        // Arrange
        using var owned = new OwnedExcelProcessScope();
        using var batch = ExcelSession.BeginBatchWithTimeouts(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(3),
            startupTimeout: ComInteropConstants.DefaultOperationTimeout,
            _testFileCopy!);

        Assert.Equal(_testFileCopy, batch.Execute((ctx, _) => ctx.Book.FullName));

        // Trigger timeout
        Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });

        // Act — measure Dispose time
        var sw = Stopwatch.StartNew();
        batch.Dispose();
        sw.Stop();

        _output.WriteLine($"Dispose after timeout completed in {sw.Elapsed.TotalSeconds:F1}s");

        // Assert — with pre-emptive kill + 10s join timeout, Dispose should be fast
        // Before the fix, this could take 45+ seconds or hang forever
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(25),
            $"REGRESSION: Dispose() took {sw.Elapsed.TotalSeconds:F1}s after timeout. " +
            "Expected < 25s with pre-emptive kill and aggressive 10s join timeout. " +
            "Before Bug 8 fix, this would hang forever.");
        owned.AssertAllExited();

        _output.WriteLine("✓ Dispose completed with aggressive timeout (pre-emptive kill working)");
    }

    /// <summary>
    /// REGRESSION TEST: Caller cancellation also triggers aggressive cleanup.
    /// The second catch(OperationCanceledException) in Execute also sets _operationTimedOut.
    /// </summary>
    [Fact]
    public void Execute_CallerCancellation_DisposeCleansUpQuickly()
    {
        // Arrange
        using var owned = new OwnedExcelProcessScope();
        using var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromMinutes(5), // Normal timeout — not the trigger
            _testFileCopy!);

        Assert.Equal(_testFileCopy, batch.Execute((ctx, _) => ctx.Book.FullName));

        using var cts = new CancellationTokenSource();

        // Start a long operation and cancel it after 2 seconds
        using var operationStarted = new ManualResetEventSlim(false);
        Exception? caughtException = null;

        var thread = new Thread(() =>
        {
            try
            {
                batch.Execute((ctx, ct) =>
                {
                    operationStarted.Set();
                    // Simulate work that respects cancellation poorly (simulates stuck COM call)
                    Thread.Sleep(TimeSpan.FromSeconds(30));
                    return 0;
                }, cts.Token);
            }
            catch (Exception ex)
            {
                caughtException = ex;
            }
        });

        thread.Start();
        Assert.True(operationStarted.Wait(TimeSpan.FromSeconds(10)));

        // Cancel from caller side
        cts.Cancel();
        Assert.True(thread.Join(TimeSpan.FromSeconds(15)));
        Assert.IsAssignableFrom<OperationCanceledException>(caughtException);
        Assert.True(batch.HasTimedOutOperation);
        var rejectedCallbackRan = false;
        var rejected = Assert.Throws<TimeoutException>(() => batch.Execute((_, _) =>
        {
            rejectedCallbackRan = true;
            return 42;
        }));
        Assert.Contains("previous operation", rejected.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(rejectedCallbackRan);

        _output.WriteLine($"Operation exception: {caughtException?.GetType().Name}: {caughtException?.Message}");

        // Act — Dispose should use aggressive cleanup since _operationTimedOut is set
        var sw = Stopwatch.StartNew();
        batch.Dispose();
        sw.Stop();

        _output.WriteLine($"Dispose after cancellation completed in {sw.Elapsed.TotalSeconds:F1}s");

        // Assert — Dispose should not hang
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(30),
            $"Dispose took {sw.Elapsed.TotalSeconds:F1}s after cancellation — expected < 30s");
        owned.AssertAllExited();

        _output.WriteLine("✓ Dispose completed after caller cancellation");
    }

    /// <summary>
    /// REGRESSION TEST: After timeout, subsequent Execute calls must throw TimeoutException
    /// immediately instead of queueing work on the stuck STA thread.
    /// Before this fix, the second caller would queue work and block until its own timeout
    /// expired — causing the entire server to appear hung for up to timeoutSeconds.
    /// </summary>
    [Fact]
    public void Execute_AfterPreviousTimeout_FailsFastWithTimeoutException()
    {
        // Arrange — short timeout to trigger the first timeout quickly
        using var owned = new OwnedExcelProcessScope();
        using var batch = ExcelSession.BeginBatchWithTimeouts(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(3),
            startupTimeout: ComInteropConstants.DefaultOperationTimeout,
            _testFileCopy!);

        // Warm up
        Assert.Equal(_testFileCopy, batch.Execute((ctx, _) => ctx.Book.FullName));

        // Trigger timeout on first operation
        Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });

        _output.WriteLine("First timeout triggered, now calling Execute again...");

        // Act — second Execute should fail FAST (not wait for its own timeout)
        var sw = Stopwatch.StartNew();
        var rejectedCallbackRan = false;
        var ex = Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) => { rejectedCallbackRan = true; return 42; });
        });
        sw.Stop();

        _output.WriteLine($"Second Execute threw in {sw.Elapsed.TotalMilliseconds:F0}ms: {ex.Message}");

        // Assert — must be near-instant, not another 3+ second timeout wait
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(1),
            $"REGRESSION: Second Execute took {sw.Elapsed.TotalSeconds:F1}s — expected < 1s. " +
            "The fail-fast pre-check for _operationTimedOut may not be working.");
        Assert.Contains("previous operation", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(rejectedCallbackRan);

        // Cleanup
        batch.Dispose();
        owned.AssertAllExited();
        _output.WriteLine("✓ Subsequent Execute after timeout fails fast");
    }

    /// <summary>
    /// Verify that ExcelProcessId is captured during session creation.
    /// This is a prerequisite for the pre-emptive kill to work.
    /// </summary>
    [Fact]
    public void BeginBatch_CapturesExcelProcessId()
    {
        // Arrange & Act
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        // Assert
        Assert.NotNull(batch.ExcelProcessId);
        Assert.True(batch.ExcelProcessId > 0, "ExcelProcessId should be a valid PID");

        // Verify the process actually exists
        using var process = Process.GetProcessById(batch.ExcelProcessId.Value);
        Assert.False(process.HasExited, "Excel process should be running");
        Assert.Equal("EXCEL", process.ProcessName, ignoreCase: true);

        _output.WriteLine($"✓ ExcelProcessId captured: {batch.ExcelProcessId}");
    }
}
