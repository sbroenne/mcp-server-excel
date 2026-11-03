using System.Diagnostics;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("RequiresExcel", "false")]
[Trait("AdapterTestKind", "System")]
[Trait("Speed", "Medium")]
public sealed class DaemonStartupCancellationTests(ITestOutputHelper output) : CliDaemonTestBase(output)
{
    [Fact]
    public async Task Startup_CancelWhileWaitingForLock_StopsPromptlyWithoutStartingDaemon()
    {
        var pipeName = $"startup-cancel-{Guid.NewGuid():N}";
        await using var holder = await HeldMutex.AcquireAsync(DaemonStartupLock.GetStartupMutexName(pipeName));
        using var cancellation = new CancellationTokenSource();
        var clock = new WaitObservingTimeProvider();
        var deadline = OperationDeadline.Start(DaemonAutoStart.StartupReadyTimeout, clock);
        var startup = DaemonAutoStart.TryStartDaemonWithStartupLockAsync(
            pipeName, deadline, cancellation.Token);

        try
        {
            var waitingThread = await clock.WaitingThread.Task.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.True(SpinWait.SpinUntil(
                () => (waitingThread.ThreadState & System.Threading.ThreadState.WaitSleepJoin) != 0,
                TimeSpan.FromSeconds(5)), "Startup did not reach the held mutex wait.");

            var elapsed = Stopwatch.StartNew();
            cancellation.Cancel();
            var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => startup.WaitAsync(TimeSpan.FromSeconds(2)));
            Assert.Equal(cancellation.Token, exception.CancellationToken);
            Assert.InRange(elapsed.Elapsed, TimeSpan.Zero, TimeSpan.FromSeconds(2));
            Assert.False(DaemonAutoStart.IsDaemonMutexHeld(pipeName));
            Assert.False(DaemonAutoStart.IsDaemonStartupInProgress(pipeName));
        }
        finally
        {
            cancellation.Cancel();
            // Keep the lock held until the cancelled attempt has actually finished.
            Assert.Same(startup, await Task.WhenAny(
                startup, Task.Delay(TimeSpan.FromSeconds(15))));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task AlreadyCancelled_DoesNotStartOrRunAction(bool startup)
    {
        var pipeName = $"startup-pre-cancel-{Guid.NewGuid():N}";
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var actionRan = false;
        Task operation = startup
            ? DaemonAutoStart.TryStartDaemonWithStartupLockAsync(
                pipeName,
                OperationDeadline.Start(DaemonAutoStart.StartupReadyTimeout),
                cancellation.Token)
            : DaemonStartupLock.WithLockAsync(
                pipeName,
                () =>
                {
                    actionRan = true;
                    return Task.FromResult(true);
                },
                cancellation.Token);

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => operation);
        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.False(actionRan);
        Assert.False(DaemonAutoStart.IsDaemonMutexHeld(pipeName));
        Assert.False(DaemonAutoStart.IsDaemonStartupInProgress(pipeName));
        Assert.True(await DaemonStartupLock.WithLockAsync(
            pipeName, () => Task.FromResult(true), CancellationToken.None));
    }

    [Fact]
    public async Task SharedLockWait_CancelWhileHeld_PreservesCancellation()
    {
        var pipeName = $"shared-lock-cancel-{Guid.NewGuid():N}";
        await using var holder = await HeldMutex.AcquireAsync(DaemonStartupLock.GetStartupMutexName(pipeName));
        using var cancellation = new CancellationTokenSource();
        var waitingThread = new TaskCompletionSource<Thread>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        var wait = Task.Factory.StartNew(() =>
        {
            using var mutex = new Mutex(false, DaemonStartupLock.GetStartupMutexName(pipeName));
            waitingThread.SetResult(Thread.CurrentThread);
            var acquired = DaemonStartupLock.TryAcquire(
                mutex, DaemonStartupLock.Timeout, cancellation.Token);
            if (acquired)
                mutex.ReleaseMutex();
            return acquired;
        }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);

        try
        {
            var thread = await waitingThread.Task.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.True(SpinWait.SpinUntil(
                () => (thread.ThreadState & System.Threading.ThreadState.WaitSleepJoin) != 0,
                TimeSpan.FromSeconds(5)), "The shared helper did not reach the held mutex wait.");

            cancellation.Cancel();
            var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => wait.WaitAsync(TimeSpan.FromSeconds(2)));
            Assert.Equal(cancellation.Token, exception.CancellationToken);
        }
        finally
        {
            cancellation.Cancel();
            Assert.Same(wait, await Task.WhenAny(
                wait, Task.Delay(TimeSpan.FromSeconds(15))));
        }
    }

    [Fact]
    public async Task SharedLockWait_ContentionTimeout_DoesNotAcquire()
    {
        var pipeName = $"shared-lock-timeout-{Guid.NewGuid():N}";
        await using var holder = await HeldMutex.AcquireAsync(DaemonStartupLock.GetStartupMutexName(pipeName));

        var elapsed = Stopwatch.StartNew();
        var acquired = await Task.Run(() =>
        {
            using var mutex = new Mutex(false, DaemonStartupLock.GetStartupMutexName(pipeName));
            var ownsMutex = DaemonStartupLock.TryAcquire(
                mutex, TimeSpan.FromMilliseconds(150), CancellationToken.None);
            if (ownsMutex)
                mutex.ReleaseMutex();
            return ownsMutex;
        }).WaitAsync(TimeSpan.FromSeconds(5));

        Assert.False(acquired);
        Assert.InRange(elapsed.Elapsed, TimeSpan.FromMilliseconds(100), TimeSpan.FromSeconds(5));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task WithLock_ActionFailure_ReleasesLockForNextCaller(bool cancel)
    {
        var pipeName = $"shared-lock-failure-{Guid.NewGuid():N}";
        using var cancellation = new CancellationTokenSource();
        Exception failure = cancel
            ? new OperationCanceledException(cancellation.Token)
            : new InvalidOperationException("Test action failed.");

        var exception = await Record.ExceptionAsync(
            () => DaemonStartupLock.WithLockAsync<bool>(
                pipeName,
                () =>
                {
                    if (cancel)
                        cancellation.Cancel();
                    throw failure;
                },
                cancellation.Token));

        Assert.Same(failure, exception);
        Assert.Equal(42, await DaemonStartupLock.WithLockAsync(
            pipeName, () => Task.FromResult(42), CancellationToken.None));
    }

    [Fact]
    public async Task WithLock_CancelledContender_DoesNotRunActionOrBlockNextCaller()
    {
        var pipeName = $"shared-lock-contender-{Guid.NewGuid():N}";
        using var cancellation = new CancellationTokenSource();
        var actionRan = false;
        Task<bool> contender;
        await using (var holder = await HeldMutex.AcquireAsync(DaemonStartupLock.GetStartupMutexName(pipeName)))
        {
            contender = DaemonStartupLock.WithLockAsync(
                pipeName,
                () =>
                {
                    actionRan = true;
                    return Task.FromResult(true);
                },
                cancellation.Token);
            cancellation.Cancel();
            var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => contender.WaitAsync(TimeSpan.FromSeconds(2)));
            Assert.Equal(cancellation.Token, exception.CancellationToken);
            Assert.False(actionRan);
        }

        Assert.Equal(42, await DaemonStartupLock.WithLockAsync(
            pipeName, () => Task.FromResult(42), CancellationToken.None));
    }

    [Fact]
    public async Task WithLock_AbandonedMutex_RunsActionAndReleasesLock()
    {
        var pipeName = $"shared-lock-abandoned-{Guid.NewGuid():N}";
        using var mutex = CreateAbandonedMutex(DaemonStartupLock.GetStartupMutexName(pipeName));

        Assert.Equal(42, await DaemonStartupLock.WithLockAsync(
            pipeName, () => Task.FromResult(42), CancellationToken.None));
        Assert.Equal(43, await DaemonStartupLock.WithLockAsync(
            pipeName, () => Task.FromResult(43), CancellationToken.None));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Startup_ExpiredDeadline_DoesNotAcquireOrStartDaemon(bool cancel)
    {
        var pipeName = $"startup-expired-{Guid.NewGuid():N}";
        using var cancellation = new CancellationTokenSource();
        if (cancel)
            cancellation.Cancel();

        var startup = DaemonAutoStart.TryStartDaemonWithStartupLockAsync(
            pipeName, OperationDeadline.Start(TimeSpan.Zero), cancellation.Token);
        if (cancel)
        {
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => startup);
        }
        else
        {
            Assert.Equal(DaemonAutoStart.StartOutcome.LockUnavailable, await startup);
        }
        Assert.False(DaemonAutoStart.IsDaemonMutexHeld(pipeName));
        Assert.False(DaemonAutoStart.IsDaemonStartupInProgress(pipeName));
    }

    private sealed class WaitObservingTimeProvider : TimeProvider
    {
        private int _timestampReads;

        internal TaskCompletionSource<Thread> WaitingThread { get; } =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        public override long GetTimestamp()
        {
            if (Interlocked.Increment(ref _timestampReads) == 2)
            {
                WaitingThread.SetResult(Thread.CurrentThread);
            }

            return base.GetTimestamp();
        }
    }
}
