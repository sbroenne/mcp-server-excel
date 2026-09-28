using Sbroenne.ExcelMcp.Service.Rpc;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "ServiceClient")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class ServiceClientDeadlineTests
{
    [Fact]
    public void GetStepTimeout_UsesOneControlledTotalBudgetAcrossSteps()
    {
        var clock = new AdvancingTimeProvider();
        var startedAt = clock.GetTimestamp();

        clock.Advance(TimeSpan.FromSeconds(3));
        Assert.Equal(
            TimeSpan.FromSeconds(5),
            ServiceClient.GetStepTimeout(
                TimeSpan.FromSeconds(5),
                TimeSpan.FromSeconds(10),
                startedAt,
                clock));

        clock.Advance(TimeSpan.FromSeconds(4));
        Assert.Equal(
            TimeSpan.FromSeconds(3),
            ServiceClient.GetStepTimeout(
                TimeSpan.FromSeconds(5),
                TimeSpan.FromSeconds(10),
                startedAt,
                clock));

        clock.Advance(TimeSpan.FromSeconds(3));
        Assert.Throws<TimeoutException>(() =>
            ServiceClient.GetStepTimeout(
                TimeSpan.FromSeconds(5),
                TimeSpan.FromSeconds(10),
                startedAt,
                clock));
    }

    [Fact]
    public async Task SendAsync_PendingConnectionUsesControlledTimeout()
    {
        var clock = new ManualTimerTimeProvider();
        using var client = new ServiceClient(
            $"excelmcp-missing-{Guid.NewGuid():N}",
            connectTimeout: TimeSpan.FromSeconds(10),
            requestTimeout: TimeSpan.FromSeconds(20),
            clock);

        var sendTask = client.SendAsync(new ServiceRequest { Command = "service.ping" });
        await clock.WaitForTimerAsync();

        clock.Advance(TimeSpan.FromSeconds(10));

        var response = await sendTask.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.False(response.Success);
        Assert.Equal("Timeout", response.ErrorCategory);
        Assert.Equal("Service connection timed out", response.ErrorMessage);
    }

    [Fact]
    public async Task SendAsync_PendingRequestUsesControlledTimeout()
    {
        var pipeName = $"excelmcp-client-timeout-{Guid.NewGuid():N}";
        var requestStarted = new TaskCompletionSource<bool>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        var responseSource = new TaskCompletionSource<ServiceResponse>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        using var host = new DaemonHost(
            _ =>
            {
                requestStarted.TrySetResult(true);
                return responseSource.Task;
            },
            static () => 0);
        var hostTask = host.RunAsync(pipeName);
        var clock = new ManualTimerTimeProvider();
        using var client = new ServiceClient(
            pipeName,
            connectTimeout: TimeSpan.FromSeconds(5),
            requestTimeout: TimeSpan.FromSeconds(10),
            clock);

        var sendTask = client.SendAsync(new ServiceRequest { Command = "sheet.list" });
        await requestStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));

        clock.Advance(TimeSpan.FromSeconds(10));

        var response = await sendTask.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.False(response.Success);
        Assert.Equal("Timeout", response.ErrorCategory);
        Assert.Equal("Service request timed out", response.ErrorMessage);

        responseSource.TrySetResult(new ServiceResponse { Success = true });
        host.RequestShutdown();
        await hostTask.WaitAsync(TimeSpan.FromSeconds(5));
    }

    [Fact]
    public async Task SendAsync_ExternalCancellationIsNotReportedAsTimeout()
    {
        var pipeName = $"excelmcp-client-cancel-{Guid.NewGuid():N}";
        var requestStarted = new TaskCompletionSource<bool>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        var responseSource = new TaskCompletionSource<ServiceResponse>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        using var host = new DaemonHost(
            _ =>
            {
                requestStarted.TrySetResult(true);
                return responseSource.Task;
            },
            static () => 0);
        var hostTask = host.RunAsync(pipeName);
        var clock = new ManualTimerTimeProvider();
        using var client = new ServiceClient(
            pipeName,
            connectTimeout: TimeSpan.FromSeconds(5),
            requestTimeout: TimeSpan.FromSeconds(10),
            clock);
        using var cancellation = new CancellationTokenSource();

        var sendTask = client.SendAsync(
            new ServiceRequest { Command = "sheet.list" },
            cancellation.Token);
        await requestStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));

        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sendTask.WaitAsync(TimeSpan.FromSeconds(5)));
        responseSource.TrySetResult(new ServiceResponse { Success = true });
        host.RequestShutdown();
        await hostTask.WaitAsync(TimeSpan.FromSeconds(5));
    }

    private class AdvancingTimeProvider : TimeProvider
    {
        protected long Timestamp;

        public override long TimestampFrequency => TimeSpan.TicksPerSecond;

        public override long GetTimestamp() => Timestamp;

        internal virtual void Advance(TimeSpan elapsed) => Timestamp += elapsed.Ticks;
    }

    private sealed class ManualTimerTimeProvider : AdvancingTimeProvider
    {
        private readonly object _gate = new();
        private readonly List<ManualTimer> _timers = [];
        private readonly TaskCompletionSource<bool> _timerCreated =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        public override ITimer CreateTimer(
            TimerCallback callback,
            object? state,
            TimeSpan dueTime,
            TimeSpan period)
        {
            var timer = new ManualTimer(this, callback, state, dueTime, period);
            lock (_gate)
            {
                _timers.Add(timer);
            }

            _timerCreated.TrySetResult(true);
            return timer;
        }

        internal async Task WaitForTimerAsync()
        {
            await _timerCreated.Task.WaitAsync(TimeSpan.FromSeconds(5));
        }

        internal override void Advance(TimeSpan elapsed)
        {
            base.Advance(elapsed);
            ManualTimer[] dueTimers;
            lock (_gate)
            {
                dueTimers = _timers
                    .Where(timer => timer.IsDue(Timestamp))
                    .ToArray();
            }

            foreach (var timer in dueTimers)
            {
                timer.Fire(Timestamp);
            }
        }

        private void Remove(ManualTimer timer)
        {
            lock (_gate)
            {
                _timers.Remove(timer);
            }
        }

        private sealed class ManualTimer(
            ManualTimerTimeProvider owner,
            TimerCallback callback,
            object? state,
            TimeSpan dueTime,
            TimeSpan period) : ITimer
        {
            private long _dueAt = owner.Timestamp + dueTime.Ticks;
            private TimeSpan _period = period;
            private bool _disposed;

            internal bool IsDue(long timestamp) => !_disposed && timestamp >= _dueAt;

            internal void Fire(long timestamp)
            {
                if (_disposed)
                {
                    return;
                }

                if (_period == Timeout.InfiniteTimeSpan)
                {
                    Dispose();
                }
                else
                {
                    _dueAt = timestamp + _period.Ticks;
                }

                callback(state);
            }

            public bool Change(TimeSpan dueTime, TimeSpan newPeriod)
            {
                if (_disposed)
                {
                    return false;
                }

                _dueAt = owner.Timestamp + dueTime.Ticks;
                _period = newPeriod;
                return true;
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                owner.Remove(this);
            }

            public ValueTask DisposeAsync()
            {
                Dispose();
                return ValueTask.CompletedTask;
            }
        }
    }
}
