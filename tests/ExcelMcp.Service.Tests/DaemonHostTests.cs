using Sbroenne.ExcelMcp.Service.Rpc;
using StreamJsonRpc;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class DaemonHostTests
{
    [Fact(Timeout = 15000)]
    public async Task Shutdown_ResponseArrivesBeforeHostDrainsOpenConnection()
    {
        var pipeName = $"excelmcp-daemon-host-{Guid.NewGuid():N}";
        using var delays = new ControlledDelay();
        DaemonHost? host = null;
        host = new DaemonHost(
            request =>
            {
                Assert.Equal("service.shutdown", request.Command);
                host.RequestShutdownAfterResponse();
                return Task.FromResult(new ServiceResponse { Success = true });
            },
            static () => 0,
            delay: delays.DelayAsync);
        using (host)
        {
            var runTask = host.RunAsync(pipeName);
            using var pipe = ServiceSecurity.CreateClient(pipeName);
            await pipe.ConnectAsync(5000);
            var proxy = JsonRpc.Attach<IExcelDaemonRpc>(pipe);
            try
            {
                var response = await proxy.ProcessCommandAsync(
                    new ServiceRequest { Command = "service.shutdown" });

                Assert.True(response.Success);
                var shutdownDelay = await delays.TakeAsync();
                Assert.Equal(TimeSpan.FromMilliseconds(100), shutdownDelay.Duration);
                shutdownDelay.Complete();
                await host.WaitForAcceptLoopStoppedAsync().WaitAsync(TimeSpan.FromSeconds(5));
                Assert.False(
                    runTask.IsCompleted,
                    "The host must drain the retained client connection after sending the shutdown response.");
            }
            finally
            {
                ((IDisposable)proxy).Dispose();
            }

            await runTask.WaitAsync(TimeSpan.FromSeconds(5));
        }
    }

    [Fact(Timeout = 15000)]
    public async Task IdleTimeout_UsesControlledClockAndStopsAcceptance()
    {
        var clock = new ManualTimeProvider();
        using var delays = new ControlledDelay();
        using var host = new DaemonHost(
            static _ => Task.FromResult(new ServiceResponse { Success = true }),
            static () => 0,
            clock,
            delays.DelayAsync);

        var runTask = host.RunAsync(
            $"excelmcp-daemon-idle-{Guid.NewGuid():N}",
            TimeSpan.FromSeconds(30));
        var idleCheck = await delays.TakeAsync();
        Assert.Equal(TimeSpan.FromSeconds(30), idleCheck.Duration);

        clock.Advance(TimeSpan.FromSeconds(30));
        idleCheck.Complete();

        await host.WaitForAcceptLoopStoppedAsync().WaitAsync(TimeSpan.FromSeconds(5));
        await runTask.WaitAsync(TimeSpan.FromSeconds(5));
    }

    [Fact(Timeout = 15000)]
    public async Task ActiveSession_RefreshesIdleActivityBeforeTimeout()
    {
        var clock = new ManualTimeProvider();
        using var delays = new ControlledDelay();
        var sessionCount = 1;
        using var host = new DaemonHost(
            static _ => Task.FromResult(new ServiceResponse { Success = true }),
            () => Volatile.Read(ref sessionCount),
            clock,
            delays.DelayAsync);
        var runTask = host.RunAsync(
            $"excelmcp-daemon-active-{Guid.NewGuid():N}",
            TimeSpan.FromSeconds(30));

        var activeCheck = await delays.TakeAsync();
        clock.Advance(TimeSpan.FromSeconds(30));
        activeCheck.Complete();

        var beforeTimeout = await delays.TakeAsync();
        Volatile.Write(ref sessionCount, 0);
        clock.Advance(TimeSpan.FromSeconds(29));
        beforeTimeout.Complete();

        var timeoutCheck = await delays.TakeAsync();
        Assert.False(runTask.IsCompleted);
        clock.Advance(TimeSpan.FromSeconds(1));
        timeoutCheck.Complete();

        await host.WaitForAcceptLoopStoppedAsync().WaitAsync(TimeSpan.FromSeconds(5));
        await runTask.WaitAsync(TimeSpan.FromSeconds(5));
    }

    [Fact(Timeout = 15000)]
    public async Task AcceptFailures_BackOffExponentiallyToMaximum()
    {
        using var delays = new ControlledDelay();
        using var host = new DaemonHost(
            static _ => Task.FromResult(new ServiceResponse { Success = true }),
            static () => 0,
            delay: delays.DelayAsync,
            serverFactory: static _ => throw new IOException("controlled accept failure"));
        var runTask = host.RunAsync($"excelmcp-daemon-backoff-{Guid.NewGuid():N}");
        var expected = new[]
        {
            100,
            200,
            400,
            800,
            1600,
            3200,
            5000,
            5000
        };

        foreach (var milliseconds in expected)
        {
            var delay = await delays.TakeAsync();
            Assert.Equal(TimeSpan.FromMilliseconds(milliseconds), delay.Duration);
            delay.Complete();
        }

        var pending = await delays.TakeAsync();
        host.RequestShutdown();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => pending.Completion);
        await runTask.WaitAsync(TimeSpan.FromSeconds(5));
    }

    private sealed class ManualTimeProvider : TimeProvider
    {
        private DateTimeOffset _utcNow = DateTimeOffset.UnixEpoch;

        public override DateTimeOffset GetUtcNow() => _utcNow;

        internal void Advance(TimeSpan duration) => _utcNow += duration;
    }

    private sealed class ControlledDelay : IDisposable
    {
        private readonly Queue<DelayRequest> _requests = new();
        private readonly SemaphoreSlim _available = new(0);
        private readonly object _gate = new();

        internal Task DelayAsync(
            TimeSpan duration,
            CancellationToken cancellationToken)
        {
            var request = new DelayRequest(duration, cancellationToken);
            lock (_gate)
            {
                _requests.Enqueue(request);
            }

            _available.Release();
            return request.Completion;
        }

        internal async Task<DelayRequest> TakeAsync()
        {
            await _available.WaitAsync(TimeSpan.FromSeconds(5));
            lock (_gate)
            {
                return _requests.Dequeue();
            }
        }

        public void Dispose() => _available.Dispose();
    }

    private sealed class DelayRequest
    {
        private readonly TaskCompletionSource _completion =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly CancellationTokenRegistration _registration;

        internal DelayRequest(
            TimeSpan duration,
            CancellationToken cancellationToken)
        {
            Duration = duration;
            _registration = cancellationToken.Register(
                static state => ((TaskCompletionSource)state!).TrySetCanceled(),
                _completion);
        }

        internal TimeSpan Duration { get; }
        internal Task Completion => _completion.Task;

        internal void Complete()
        {
            _registration.Dispose();
            _completion.TrySetResult();
        }
    }
}
