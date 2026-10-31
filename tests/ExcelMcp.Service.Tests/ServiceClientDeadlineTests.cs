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
    public void ConnectTimeoutConversion_PreservesInfiniteAndPositiveRounding()
    {
        Assert.Equal(Timeout.Infinite, ServiceClient.ToTimeoutMilliseconds(Timeout.InfiniteTimeSpan));
        Assert.Equal(1, ServiceClient.ToTimeoutMilliseconds(TimeSpan.FromTicks(1)));
        Assert.Equal(2, ServiceClient.ToTimeoutMilliseconds(TimeSpan.FromMilliseconds(1.1)));
        Assert.Equal(6000, ServiceClient.ToTimeoutMilliseconds(TimeSpan.FromSeconds(6)));
    }
    [Fact]
    public void ControlledTimer_InfiniteDueTimeRemainsDisabledUntilChanged()
    {
        var clock = new ManualTimerTimeProvider();
        int calls = 0;
        using var timer = clock.CreateTimer(_ => calls++, null, Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
        clock.Advance(TimeSpan.FromHours(1));
        Assert.Equal(0, calls);
        Assert.True(timer.Change(TimeSpan.FromSeconds(5), Timeout.InfiniteTimeSpan));
        clock.Advance(TimeSpan.FromSeconds(4));
        Assert.Equal(0, calls);
        clock.Advance(TimeSpan.FromSeconds(1));
        Assert.Equal(1, calls);
        clock.Advance(TimeSpan.FromHours(1));
        Assert.Equal(1, calls);
    }

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
    public void GetStepTimeout_InfiniteStepStillHonorsFiniteTotalBudget()
    {
        var clock = new AdvancingTimeProvider();
        var startedAt = clock.GetTimestamp();
        clock.Advance(TimeSpan.FromSeconds(4));
        Assert.Equal(TimeSpan.FromSeconds(6), ServiceClient.GetStepTimeout(
            Timeout.InfiniteTimeSpan, TimeSpan.FromSeconds(10), startedAt, clock));
        Assert.Equal(Timeout.InfiniteTimeSpan, ServiceClient.GetStepTimeout(
            Timeout.InfiniteTimeSpan, null, startedAt, clock));
        clock.Advance(TimeSpan.FromSeconds(6));
        Assert.Throws<TimeoutException>(() => ServiceClient.GetStepTimeout(
            Timeout.InfiniteTimeSpan, TimeSpan.FromSeconds(10), startedAt, clock));
    }

    [Fact]
    public async Task SendAsync_TotalBudgetExpiresBeforeConnection_ReturnsContextualTimeout()
    {
        using var client = new ServiceClient(
            $"excelmcp-expired-{Guid.NewGuid():N}",
            connectTimeout: TimeSpan.FromSeconds(5),
            requestTimeout: TimeSpan.FromSeconds(10),
            new ExpiringBeforeConnectTimeProvider());
        var request = new ServiceRequest
        {
            Command = "service.ping",
            SessionId = "retained-request-context"
        };

        var response = await client.SendAsync(request, TimeSpan.FromSeconds(1));

        Assert.False(response.Success);
        Assert.Equal(request.Command, response.Command);
        Assert.Equal(request.SessionId, response.SessionId);
        Assert.Equal("Timeout", response.ErrorCategory);
        Assert.Equal(nameof(TimeoutException), response.ExceptionType);
        Assert.Equal("Service connection timed out", response.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(response.Result));
    }

    [Fact]
    public async Task SendAsync_CancelledCallerWithExpiredBudget_PreservesCancellation()
    {
        using var client = new ServiceClient(
            $"excelmcp-cancelled-expired-{Guid.NewGuid():N}",
            connectTimeout: TimeSpan.FromSeconds(5),
            requestTimeout: TimeSpan.FromSeconds(10),
            new ExpiringBeforeConnectTimeProvider());
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            client.SendAsync(
                new ServiceRequest { Command = "service.ping" },
                TimeSpan.FromSeconds(1),
                cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
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
        var abandoned = new TaskCompletionSource<ServiceRequest>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        using var host = new DaemonHost(
            _ =>
            {
                requestStarted.TrySetResult(true);
                return responseSource.Task;
            },
            requestAbandoned: (request, _) => abandoned.TrySetResult(request));
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
        Assert.Equal("sheet.list",
            (await abandoned.Task.WaitAsync(TimeSpan.FromSeconds(5))).Command);
        Assert.False(responseSource.Task.IsCompleted);

        responseSource.TrySetResult(new ServiceResponse { Success = true });
        host.RequestShutdown();
        await hostTask.WaitAsync(TimeSpan.FromSeconds(5));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SendAsync_ExternalCancellationIsNotReportedAsTimeout(bool waitWithoutDeadline)
    {
        var pipeName = $"excelmcp-client-cancel-{Guid.NewGuid():N}";
        var requestStarted = new TaskCompletionSource<bool>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        var responseSource = new TaskCompletionSource<ServiceResponse>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        var abandoned = new TaskCompletionSource<ServiceRequest>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        using var host = new DaemonHost(
            _ =>
            {
                requestStarted.TrySetResult(true);
                return responseSource.Task;
            },
            requestAbandoned: (request, _) => abandoned.TrySetResult(request));
        var hostTask = host.RunAsync(pipeName);
        var clock = new ManualTimerTimeProvider();
        using var client = new ServiceClient(
            pipeName,
            connectTimeout: TimeSpan.FromSeconds(5),
            requestTimeout: waitWithoutDeadline ? Timeout.InfiniteTimeSpan : TimeSpan.FromSeconds(10),
            clock);
        using var cancellation = new CancellationTokenSource();

        var sendTask = client.SendAsync(
            new ServiceRequest { Command = "sheet.list" },
            cancellation.Token);
        await requestStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));

        if (waitWithoutDeadline)
        {
            clock.Advance(TimeSpan.FromSeconds(65));
            Assert.False(sendTask.IsCompleted);
        }

        cancellation.Cancel();

        var error = await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sendTask.WaitAsync(TimeSpan.FromSeconds(5)));
        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Equal("sheet.list",
            (await abandoned.Task.WaitAsync(TimeSpan.FromSeconds(5))).Command);
        Assert.False(responseSource.Task.IsCompleted);
        responseSource.TrySetResult(new ServiceResponse { Success = true });
        host.RequestShutdown();
        await hostTask.WaitAsync(TimeSpan.FromSeconds(5));
    }

    [Fact]
    public async Task SendAsync_CallerCancellationWinsTransportLossBeforeLinkedCancellationRuns()
    {
        var pipeName = $"excelmcp-client-cancel-disconnect-{Guid.NewGuid():N}";
        var transportCreated = new TaskCompletionSource<System.IO.Pipes.NamedPipeServerStream>(TaskCreationOptions.RunContinuationsAsynchronously);
        var requestStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var pendingResponse = new TaskCompletionSource<ServiceResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var releaseCancellation = new ManualResetEventSlim();
        using var host = new DaemonHost(_ =>
        {
            requestStarted.TrySetResult();
            return pendingResponse.Task;
        }, serverFactory: name =>
        {
            var pipe = ServiceSecurity.CreateSecureServer(name);
            transportCreated.TrySetResult(pipe);
            return pipe;
        });
        var hostTask = host.RunAsync(pipeName);
        using var client = new ServiceClient(pipeName, requestTimeout: Timeout.InfiniteTimeSpan);
        using var cancellation = new CancellationTokenSource();
        var send = client.SendAsync(new ServiceRequest { Command = "sheet.list" }, cancellation.Token);
        await requestStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        var transport = await transportCreated.Task.WaitAsync(TimeSpan.FromSeconds(5));
        // Cancellation callbacks run in reverse registration order. Hold propagation to
        // the client's linked token after making the original caller token cancelled.
        using var registration = cancellation.Token.Register(() =>
        {
            transport.Dispose();
            if (!releaseCancellation.Wait(TimeSpan.FromSeconds(10)))
                throw new TimeoutException("The cancellation/disconnect test was not released.");
        });
        var cancelling = Task.Run(cancellation.Cancel);
        try
        {
            var error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => send.WaitAsync(TimeSpan.FromSeconds(5)));
            Assert.Equal(cancellation.Token, error.CancellationToken);
            Assert.False(pendingResponse.Task.IsCompleted);
        }
        finally
        {
            releaseCancellation.Set();
            await cancelling.WaitAsync(TimeSpan.FromSeconds(5));
            pendingResponse.TrySetResult(new ServiceResponse { Success = true });
            host.RequestShutdown();
            await hostTask.WaitAsync(TimeSpan.FromSeconds(5));
        }
    }

    [Theory]
    [InlineData("session.open", false, true, true, 0)]
    [InlineData("session.open", true, true, true, 0)]
    [InlineData("session.create", false, true, true, 0)]
    [InlineData("session.create", true, true, true, 0)]
    [InlineData("session.open", false, false, true, 0)]
    [InlineData("session.create", true, false, true, 0)]
    [InlineData("session.open", true, true, false, 0)]
    [InlineData("sheet.list", true, true, true, 0)]
    [InlineData("session.open", false, true, true, 1)]
    [InlineData("session.create", true, true, true, 1)]
    [InlineData("session.open", false, true, true, 2)]
    [InlineData("session.create", true, true, true, 2)]
    public async Task SendAsync_CompletedSessionCancelledBeforeReturn_RevealsOnlySuccessfulSessionCreation(
        string command, bool useControlEndpoint, bool responseHasSessionId, bool success, int recoveryFailure)
    {
        var pipeName = $"excelmcp-completed-cancel-{Guid.NewGuid():N}";
        const string sessionId = "completed-session";
        var requests = new System.Collections.Concurrent.ConcurrentQueue<ServiceRequest>();
        var revealed = false;
        var controlRequests = new System.Collections.Concurrent.ConcurrentQueue<ServiceRequest>();
        Task<ServiceResponse> Handle(ServiceRequest request)
        {
            requests.Enqueue(request);
            if (request.Command == "window.show")
            {
                Assert.Equal(sessionId, request.SessionId);
                if (recoveryFailure == 2)
                    throw new InvalidOperationException("The recovery endpoint failed.");
                revealed = recoveryFailure == 0;
            }
            return Task.FromResult(new ServiceResponse
            {
                Success = request.Command == "window.show" ? recoveryFailure == 0 : success,
                ErrorMessage = request.Command == "window.show" && recoveryFailure == 1
                    ? "The recovery request was refused." : null,
                SessionId = responseHasSessionId ? sessionId : null,
                Result = "{\"sessionId\":\"completed-session\"}"
            });
        }
        using var host = new DaemonHost(Handle);
        using var controlHost = new DaemonHost(request =>
        {
            controlRequests.Enqueue(request);
            return Handle(request);
        });
        var hostTask = host.RunAsync(pipeName);
        var controlTask = controlHost.RunAsync(ServiceClient.GetControlPipeName(pipeName));
        using var cancellation = new CancellationTokenSource();
        var clock = new CancelOnDisconnectMonitorDisposalTimeProvider(cancellation);
        using var client = new ServiceClient(pipeName, TimeSpan.FromSeconds(5),
            TimeSpan.FromSeconds(10), clock, useControlEndpoint);
        try
        {
            var error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                client.SendAsync(new ServiceRequest { Command = command }, cancellation.Token)
                    .WaitAsync(TimeSpan.FromSeconds(10)));
            Assert.Equal(cancellation.Token, error.CancellationToken);
            var shouldReveal = success && command is "session.open" or "session.create";
            Assert.Equal(shouldReveal && recoveryFailure == 0, revealed);
            Assert.Equal(shouldReveal ? new[] { command, "window.show" } : new[] { command },
                requests.Select(request => request.Command));
            Assert.Equal(shouldReveal && useControlEndpoint ? 1 : 0, controlRequests.Count);
            Assert.True(clock.CancelledAfterResponse);
        }
        finally
        {
            host.RequestShutdown();
            controlHost.RequestShutdown();
            await Task.WhenAll(hostTask, controlTask).WaitAsync(TimeSpan.FromSeconds(5));
        }
    }

    // The disconnect poll is cancelled only after the RPC reply completed. Cancel
    // the caller at that boundary, without sleeps or racing the server handler.
    private sealed class CancelOnDisconnectMonitorDisposalTimeProvider(CancellationTokenSource cancellation) : TimeProvider
    {
        public bool CancelledAfterResponse { get; private set; }

        public override ITimer CreateTimer(TimerCallback callback, object? state, TimeSpan dueTime, TimeSpan period)
        {
            if (dueTime != TimeSpan.FromMilliseconds(250))
                return System.CreateTimer(callback, state, dueTime, period);
            return new DisposalTimer(() =>
            {
                CancelledAfterResponse = true;
                cancellation.Cancel();
            });
        }

        private sealed class DisposalTimer(Action onDispose) : ITimer
        {
            private int _disposed;
            public bool Change(TimeSpan dueTime, TimeSpan period) => true;
            public void Dispose()
            {
                if (Interlocked.Exchange(ref _disposed, 1) == 0)
                    onDispose();
            }
            public ValueTask DisposeAsync()
            {
                Dispose();
                return ValueTask.CompletedTask;
            }
        }
    }

    private sealed class ExpiringBeforeConnectTimeProvider : TimeProvider
    {
        private int _timestampReads;

        public override long TimestampFrequency => TimeSpan.TicksPerSecond;

        public override long GetTimestamp() =>
            Interlocked.Increment(ref _timestampReads) == 1
                ? 0
                : TimeSpan.FromSeconds(1).Ticks;
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
            private bool _enabled = dueTime != Timeout.InfiniteTimeSpan;
            private TimeSpan _period = period;
            private bool _disposed;

            internal bool IsDue(long timestamp) => !_disposed && _enabled && timestamp >= _dueAt;

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
                _enabled = dueTime != Timeout.InfiniteTimeSpan;
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
