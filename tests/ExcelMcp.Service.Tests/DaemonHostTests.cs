using System.Text.Json;
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
    private static readonly string[] PressureProbeCommands =
        ["service.ping", "service.status", "session.list", "range.get-values"];

    [Fact(Timeout = 15000)]
    public async Task RequestsAtCapacity_ControlRequestsRemainAvailable()
    {
        var pipeName = $"excelmcp-daemon-control-{Guid.NewGuid():N}";
        var entered = new SemaphoreSlim(0);
        var release = new TaskCompletionSource<ServiceResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var host = new DaemonHost(request =>
        {
            if (request.SessionId is not null)
            {
                entered.Release();
                return release.Task;
            }
            return Task.FromResult(new ServiceResponse { Success = true });
        });
        var run = host.RunAsync(pipeName);
        var requests = Enumerable.Range(0, 10).Select(async index =>
        {
            using var client = new ServiceClient(pipeName);
            return await client.SendAsync(new ServiceRequest
            {
                Command = "session.save",
                SessionId = $"session-{index}"
            });
        }).ToArray();
        var check = await Record.ExceptionAsync(async () =>
        {
            for (var index = 0; index < 10; index++)
                Assert.True(await entered.WaitAsync(TimeSpan.FromSeconds(2)));
            foreach (var command in PressureProbeCommands.Take(3))
            {
                using var client = new ServiceClient(pipeName);
                Assert.True((await client.SendAsync(new ServiceRequest { Command = command },
                    TimeSpan.FromSeconds(2))).Success);
            }
            using var overflow = new ServiceClient(pipeName);
            var rejected = await overflow.SendAsync(new ServiceRequest
            {
                Command = "session.save",
                SessionId = "not-started"
            }, TimeSpan.FromSeconds(2));
            Assert.False(rejected.Success);
            Assert.Equal("Busy", rejected.ErrorCategory);
            Assert.Equal("not-started", rejected.SessionId);
        });
        release.SetResult(new ServiceResponse { Success = true });
        await Task.WhenAll(requests);
        using var retry = new ServiceClient(pipeName);
        Assert.True((await retry.SendAsync(new ServiceRequest
        {
            Command = "session.save",
            SessionId = "not-started"
        }, TimeSpan.FromSeconds(2))).Success);
        host.RequestShutdown();
        await run.WaitAsync(TimeSpan.FromSeconds(5));
        entered.Dispose();
        Assert.Null(check);
    }

    [Fact(Timeout = 15000)]
    public async Task ConnectionCapacity_RefusesExcessAndRecoversAfterDisconnect()
    {
        var pipeName = $"excelmcp-daemon-connections-{Guid.NewGuid():N}";
        using var host = new DaemonHost(
            static _ => Task.FromResult(new ServiceResponse { Success = true }));
        var run = host.RunAsync(pipeName);
        var connections = new List<IDisposable>();
        try
        {
            for (var index = 0; index < DaemonHost.MaximumConnections; index++)
            {
                var pipe = ServiceSecurity.CreateClient(pipeName);
                connections.Add(pipe);
                await pipe.ConnectAsync(2000);
                var proxy = JsonRpc.Attach<IExcelDaemonRpc>(pipe);
                connections.Add((IDisposable)proxy);
                Assert.True((await proxy.ProcessCommandAsync(new ServiceRequest
                {
                    Command = "service.ping"
                }).WaitAsync(TimeSpan.FromSeconds(2))).Success);
            }
            using var client = new ServiceClient(pipeName);
            var rejected = await client.SendAsync(new ServiceRequest { Command = "service.ping" },
                TimeSpan.FromSeconds(2));
            Assert.False(rejected.Success);
            Assert.Equal("ServiceUnavailable", rejected.ErrorCategory);
            foreach (var connection in connections)
                connection.Dispose();
            connections.Clear();
            var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(3);
            ServiceResponse response;
            do
            {
                response = await client.SendAsync(new ServiceRequest { Command = "service.ping" },
                    TimeSpan.FromSeconds(1));
                if (!response.Success)
                    await Task.Delay(20);
            } while (!response.Success && DateTime.UtcNow < deadline);
            Assert.True(response.Success, response.ErrorMessage);
        }
        finally
        {
            foreach (var connection in connections)
                connection.Dispose();
            host.RequestShutdown();
            await run.WaitAsync(TimeSpan.FromSeconds(5));
        }
    }

    [Fact(Timeout = 15000)]
    public async Task HandlerFailure_ReleasesSessionForNextRequest()
    {
        var pipeName = $"excelmcp-daemon-failure-{Guid.NewGuid():N}";
        var executions = 0;
        using var host = new DaemonHost(_ =>
        {
            if (Interlocked.Increment(ref executions) == 1)
                throw new InvalidOperationException("controlled handler failure");
            return Task.FromResult(new ServiceResponse { Success = true });
        });
        var run = host.RunAsync(pipeName);
        try
        {
            using var pipe = ServiceSecurity.CreateClient(pipeName);
            await pipe.ConnectAsync(2000);
            var proxy = JsonRpc.Attach<IExcelDaemonRpc>(pipe);
            using var connection = (IDisposable)proxy;
            var request = new ServiceRequest { Command = "session.save", SessionId = "retry" };
            var failure = await Assert.ThrowsAsync<RemoteInvocationException>(() => proxy.ProcessCommandAsync(request));
            Assert.Contains("controlled handler failure", failure.Message);
            Assert.True((await proxy.ProcessCommandAsync(request)).Success);
            Assert.Equal(2, executions);
        }
        finally
        {
            host.RequestShutdown();
            await run.WaitAsync(TimeSpan.FromSeconds(5));
        }
    }

    [Fact(Timeout = 15000)]
    public async Task ExpiredCallerWait_RetainsSessionClaimUntilHandlerReturns()
    {
        var pipeName = $"excelmcp-daemon-expired-{Guid.NewGuid():N}";
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource<ServiceResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        var executions = 0;
        using var host = new DaemonHost(request =>
        {
            if (request.SessionId == "retained")
            {
                Interlocked.Increment(ref executions);
                entered.TrySetResult();
                return release.Task;
            }
            return Task.FromResult(new ServiceResponse { Success = true });
        });
        var run = host.RunAsync(pipeName);
        try
        {
            using var client = new ServiceClient(pipeName);
            var request = new ServiceRequest { Command = "session.save", SessionId = "retained" };
            var original = client.SendAsync(request, TimeSpan.FromSeconds(1));
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(3));
            var expired = await original;
            Assert.False(expired.Success);
            Assert.Equal("Timeout", expired.ErrorCategory);
            var repeat = await client.SendAsync(request, TimeSpan.FromSeconds(2));
            Assert.False(repeat.Success);
            Assert.Equal("Busy", repeat.ErrorCategory);
            Assert.Equal(1, executions);
            Assert.True((await client.SendAsync(new ServiceRequest { Command = "service.ping" },
                TimeSpan.FromSeconds(2))).Success);
        }
        finally
        {
            release.TrySetResult(new ServiceResponse { Success = true });
            host.RequestShutdown();
            await run.WaitAsync(TimeSpan.FromSeconds(5));
        }
    }

    [Fact(Timeout = 30000)]
    public async Task BlockedSession_RepeatedRequestsLeaveOtherSessionsAndHealthAvailable()
    {
        var pipeName = $"excelmcp-daemon-pressure-{Guid.NewGuid():N}";
        using var release = new ManualResetEventSlim();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var executions = 0;
        using var host = new DaemonHost(request =>
        {
            if (request.SessionId == "blocked")
            {
                Interlocked.Increment(ref executions);
                entered.TrySetResult();
                Assert.True(release.Wait(TimeSpan.FromSeconds(20)));
            }
            return Task.FromResult(new ServiceResponse { Success = true, Result = request.Args });
        });
        var run = host.RunAsync(pipeName);
        using var originalClient = new ServiceClient(pipeName);
        var original = originalClient.SendAsync(new ServiceRequest
        {
            Command = "session.save",
            SessionId = "blocked"
        });
        var check = await Record.ExceptionAsync(async () =>
        {
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(5));
            var repeated = Enumerable.Range(0, 20).Select(async _ =>
            {
                using var client = new ServiceClient(pipeName);
                var response = await client.SendAsync(new ServiceRequest
                {
                    Command = "session.save",
                    SessionId = "blocked"
                }, TimeSpan.FromSeconds(2));
                Assert.False(response.Success);
                Assert.Equal("Busy", response.ErrorCategory);
            }).ToArray();
            await Task.WhenAll(repeated);
            foreach (var command in PressureProbeCommands)
            {
                using var client = new ServiceClient(pipeName);
                var response = await client.SendAsync(new ServiceRequest
                {
                    Command = command,
                    SessionId = command.StartsWith("range.", StringComparison.Ordinal) ? "healthy" : null,
                    Args = """{"marker":"unchanged"}"""
                }, TimeSpan.FromSeconds(2));
                Assert.True(response.Success, response.ErrorMessage);
                Assert.Equal("""{"marker":"unchanged"}""", response.Result);
            }
            Assert.Equal(1, Volatile.Read(ref executions));
        });
        release.Set();
        var completion = await Record.ExceptionAsync(async () =>
            Assert.True((await original.WaitAsync(TimeSpan.FromSeconds(5))).Success));
        host.RequestShutdown();
        var shutdown = await Record.ExceptionAsync(() => run.WaitAsync(TimeSpan.FromSeconds(5)));
        Assert.All(new[] { check, completion, shutdown }, failure => Assert.Null(failure));
    }

    [Fact(Timeout = 15000)]
    public async Task Disconnect_NotifiesAbandonmentWithoutCancellingOriginalHandler()
    {
        var pipeName = $"excelmcp-daemon-disconnect-{Guid.NewGuid():N}";
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var original = new TaskCompletionSource<ServiceResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        var abandoned = new TaskCompletionSource<ServiceRequest>(TaskCreationOptions.RunContinuationsAsynchronously);
        var completed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using var host = new DaemonHost(async _ =>
        {
            entered.TrySetResult();
            var response = await original.Task;
            completed.TrySetResult();
            return response;
        }, requestAbandoned: (request, _) => abandoned.TrySetResult(request));
        var run = host.RunAsync(pipeName);
        var check = await Record.ExceptionAsync(async () =>
        {
            using var pipe = ServiceSecurity.CreateClient(pipeName);
            await pipe.ConnectAsync(5000);
            var proxy = JsonRpc.Attach<IExcelDaemonRpc>(pipe);
            using var connection = (IDisposable)proxy;
            var invocation = proxy.ProcessCommandAsync(
                new ServiceRequest { Command = "session.save", SessionId = "retained-session" });
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(5));
            connection.Dispose();
            await Assert.ThrowsAsync<ConnectionLostException>(() => invocation);
            Assert.Equal("retained-session",
                (await abandoned.Task.WaitAsync(TimeSpan.FromSeconds(5))).SessionId);
            Assert.False(original.Task.IsCompleted);
            Assert.False(completed.Task.IsCompleted);
            original.SetResult(new ServiceResponse { Success = true });
            await completed.Task.WaitAsync(TimeSpan.FromSeconds(5));
        });
        original.TrySetResult(new ServiceResponse { Success = true });
        host.RequestShutdown();
        var shutdown = await Record.ExceptionAsync(() => run.WaitAsync(TimeSpan.FromSeconds(5)));
        Assert.All(new[] { check, shutdown }, failure => Assert.Null(failure));
    }

    [Fact(Timeout = 15000)]
    public async Task MutationFailure_PartialStateSurvivesDaemonTransport()
    {
        var pipeName = $"excelmcp-daemon-state-{Guid.NewGuid():N}";
        using var state = JsonDocument.Parse(
            """{"queryName":"Query","queryCreated":true,"queryExists":true,"loadMode":"load-to-table","targetSheet":"Actual","isLoadedToDataModel":false}""");
        using var host = new DaemonHost(
            static _ => Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "Privacy",
                ErrorMessage = "Privacy failure with retained query",
                PartialState = JsonSerializer.Serialize(new
                {
                    queryName = "Query",
                    queryCreated = true,
                    queryExists = true,
                    loadMode = "load-to-table",
                    targetSheet = "Actual",
                    isLoadedToDataModel = false
                })
            }));
        var runTask = host.RunAsync(pipeName);
        try
        {
            using var client = new ServiceClient(pipeName);
            var response = await client.SendAsync(new ServiceRequest { Command = "powerquery.create" });

            Assert.False(response.Success);
            Assert.Equal("Privacy", response.ErrorCategory);
            Assert.Equal("Privacy failure with retained query", response.ErrorMessage);
            Assert.Equal(state.RootElement.GetRawText(),
                JsonSerializer.Deserialize<JsonElement>(Assert.IsType<string>(response.PartialState)).GetRawText());
        }
        finally
        {
            host.RequestShutdown();
            await runTask.WaitAsync(TimeSpan.FromSeconds(5));
        }
    }

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
                Assert.IsType<DaemonHost>(host).RequestShutdownAfterResponse();
                return Task.FromResult(new ServiceResponse { Success = true });
            },
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
    public async Task SessionlessHost_RemainsAvailableUntilExplicitShutdown()
    {
        using var host = new DaemonHost(
            static _ => Task.FromResult(new ServiceResponse { Success = true }));
        var pipeName = $"excelmcp-daemon-sessionless-{Guid.NewGuid():N}";
        var runTask = host.RunAsync(pipeName);
        try
        {
            for (var request = 0; request < 3; request++)
            {
                using var client = new ServiceClient(pipeName);
                Assert.True((await client.SendAsync(new ServiceRequest { Command = "service.ping" },
                    TimeSpan.FromSeconds(2))).Success);
                Assert.False(runTask.IsCompleted);
            }
        }
        finally
        {
            host.RequestShutdown();
            await runTask.WaitAsync(TimeSpan.FromSeconds(5));
        }
    }

    [Fact(Timeout = 15000)]
    public async Task AcceptFailures_BackOffExponentiallyToMaximum()
    {
        using var delays = new ControlledDelay();
        using var host = new DaemonHost(
            static _ => Task.FromResult(new ServiceResponse { Success = true }),
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
