using System.Diagnostics;
using Microsoft.ApplicationInsights.DataContracts;
using Sbroenne.ExcelMcp.CLI.Telemetry;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

// These tests wait on timeouts and blocked fakes, so they must not overlap other collections.
[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "Telemetry")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class DeferredTelemetrySinkTests
{
    // Timeouts under test are short; the bound is generous so a loaded machine does not
    // make the tests flaky, while still far below the 5 s / 2 s production budgets.
    private static readonly TimeSpan ShortTimeout = TimeSpan.FromMilliseconds(100);

    // Initialization budget for the tests that check it is shared and measured from Start.
    private static readonly TimeSpan SharedBudget = TimeSpan.FromMilliseconds(500);

    // Long enough for the fake flush (about 20 ms) to finish and Dispose to start before the bound expires.
    private static readonly TimeSpan ShutdownBudget = TimeSpan.FromMilliseconds(500);
    private static readonly TimeSpan MaxElapsed = TimeSpan.FromSeconds(2);

    // Fail-safe so a blocked fake can never hang the test run.
    private static readonly TimeSpan BlockedFakeLimit = TimeSpan.FromSeconds(30);

    [Fact]
    public void Track_StartsFactoryOnceAndForwardsItems()
    {
        var factoryCalls = 0;
        var fake = new FakeSink();
        var sink = CreateSink(() =>
        {
            Interlocked.Increment(ref factoryCalls);
            return fake;
        });
        var first = NewItems("sheet/list");
        var second = NewItems("range/get-values");
        var third = NewItems("session/open");

        sink.Start();
        sink.Start();
        sink.Track(first.Event, first.Request);
        sink.Track(second.Event, second.Request);
        sink.Track(third.Event, third.Request);

        Assert.Equal(1, factoryCalls);
        Assert.Equal(
            [first, second, third],
            fake.Tracked);
    }

    [Fact]
    public void Start_CalledConcurrently_RunsFactoryOnce()
    {
        using var gate = new ManualResetEventSlim();
        var factoryCalls = 0;
        var fake = new FakeSink();
        var sink = CreateSink(() =>
        {
            Interlocked.Increment(ref factoryCalls);
            gate.Wait(BlockedFakeLimit);
            return fake;
        });

        Parallel.For(0, 16, _ => sink.Start());
        gate.Set();
        var items = NewItems();
        sink.Track(items.Event, items.Request);

        Assert.Equal(1, factoryCalls);
        Assert.Equal([items], fake.Tracked);
    }

    [Fact]
    public void Shutdown_NeverStarted_DoesNotRunFactory()
    {
        var factoryCalls = 0;
        var sink = CreateSink(() =>
        {
            Interlocked.Increment(ref factoryCalls);
            return new FakeSink();
        });

        var elapsed = Measure(sink.Shutdown);

        Assert.True(elapsed < MaxElapsed, $"Shutdown took {elapsed}.");
        Assert.Equal(0, factoryCalls);
    }

    [Fact]
    public void Shutdown_NothingTracked_ReturnsPromptlyWhileFactoryIsStillBlocked()
    {
        using var gate = new ManualResetEventSlim();
        using var factoryReturned = new ManualResetEventSlim();
        var fake = new FakeSink();
        var sink = CreateSink(() =>
        {
            gate.Wait(BlockedFakeLimit);
            factoryReturned.Set();
            return fake;
        });
        sink.Start();

        try
        {
            var elapsed = Measure(sink.Shutdown);

            Assert.True(elapsed < MaxElapsed, $"Shutdown took {elapsed}.");
        }
        finally
        {
            gate.Set();
        }

        Assert.True(factoryReturned.Wait(MaxElapsed));
        Assert.Empty(fake.Calls);
    }

    [Fact]
    public void Shutdown_AfterTracking_FlushesThenDisposes()
    {
        var fake = new FakeSink();
        var sink = CreateSink(() => fake);
        var items = NewItems();
        sink.Track(items.Event, items.Request);

        sink.Shutdown();

        // The default fake flush completes asynchronously, so Dispose before
        // "flush-completed" would mean the sink was disposed mid-flush.
        Assert.Equal(["track", "flush", "flush-completed", "dispose"], fake.Calls);
    }

    [Fact]
    public void Shutdown_CalledTwice_FlushesAndDisposesOnce()
    {
        var fake = new FakeSink();
        var sink = CreateSink(() => fake);
        var items = NewItems();
        sink.Track(items.Event, items.Request);

        sink.Shutdown();
        var exception = Record.Exception(sink.Shutdown);

        Assert.Null(exception);
        Assert.Equal(["track", "flush", "flush-completed", "dispose"], fake.Calls);
    }

    [Fact]
    public void Shutdown_FlushNeverCompletes_IsBoundedByShutdownTimeout()
    {
        var fake = new FakeSink
        {
            FlushBehavior = _ => new TaskCompletionSource().Task
        };
        var sink = CreateSink(() => fake, shutdownTimeout: ShutdownBudget);
        var items = NewItems();
        sink.Track(items.Event, items.Request);

        var elapsed = Measure(sink.Shutdown);

        Assert.True(elapsed < MaxElapsed, $"Shutdown took {elapsed}.");
        Assert.Contains("flush", fake.Calls);
    }

    [Fact]
    public void Shutdown_DisposeBlocks_IsBoundedByShutdownTimeout()
    {
        using var release = new ManualResetEventSlim();
        using var disposeEntered = new ManualResetEventSlim();
        var fake = new FakeSink
        {
            OnDispose = () =>
            {
                disposeEntered.Set();
                release.Wait(BlockedFakeLimit);
            }
        };
        var sink = CreateSink(() => fake, shutdownTimeout: ShutdownBudget);
        var items = NewItems();
        sink.Track(items.Event, items.Request);

        try
        {
            var elapsed = Measure(sink.Shutdown);

            Assert.True(elapsed < MaxElapsed, $"Shutdown took {elapsed}.");
            Assert.True(disposeEntered.IsSet);
        }
        finally
        {
            release.Set();
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Shutdown_FlushFails_StillDisposesWithoutThrowing(bool failSynchronously)
    {
        var fake = new FakeSink
        {
            FlushBehavior = _ => failSynchronously
                ? throw new InvalidOperationException("flush failed")
                : Task.FromException(new InvalidOperationException("flush failed"))
        };
        var sink = CreateSink(() => fake);
        var items = NewItems();
        sink.Track(items.Event, items.Request);

        var exception = Record.Exception(sink.Shutdown);

        Assert.Null(exception);
        Assert.Equal(["track", "flush", "dispose"], fake.Calls);
    }

    [Fact]
    public void Shutdown_DisposeThrows_DoesNotThrow()
    {
        var fake = new FakeSink
        {
            OnDispose = () => throw new InvalidOperationException("dispose failed")
        };
        var sink = CreateSink(() => fake);
        var items = NewItems();
        sink.Track(items.Event, items.Request);

        var exception = Record.Exception(sink.Shutdown);

        Assert.Null(exception);
    }

    [Fact]
    public void Track_InitializationExceedsTimeout_DropsItem()
    {
        using var gate = new ManualResetEventSlim();
        using var factoryReturned = new ManualResetEventSlim();
        var fake = new FakeSink();
        var sink = CreateSink(
            () =>
            {
                gate.Wait(BlockedFakeLimit);
                factoryReturned.Set();
                return fake;
            },
            initializationTimeout: ShortTimeout);
        var items = NewItems();

        try
        {
            var elapsed = Measure(() => sink.Track(items.Event, items.Request));

            Assert.True(elapsed < MaxElapsed, $"Track took {elapsed}.");
        }
        finally
        {
            gate.Set();
        }

        Assert.True(factoryReturned.Wait(MaxElapsed));
        sink.Shutdown();
        Assert.Empty(fake.Calls);
    }

    [Fact]
    public void Track_InitializationNeverCompletes_SharesOneWaitBudgetAcrossCalls()
    {
        using var gate = new ManualResetEventSlim();
        using var factoryReturned = new ManualResetEventSlim();
        var sink = CreateSink(
            () =>
            {
                gate.Wait(BlockedFakeLimit);
                factoryReturned.Set();
                return new FakeSink();
            },
            initializationTimeout: SharedBudget);
        var items = NewItems();

        try
        {
            var elapsed = Measure(() =>
            {
                for (var i = 0; i < 4; i++)
                {
                    sink.Track(items.Event, items.Request);
                }
            });

            // One budget shared by all four calls is about 500 ms; a fresh wait per call would be about 2 s.
            Assert.True(elapsed < TimeSpan.FromMilliseconds(1200), $"Four Track calls took {elapsed}.");
        }
        finally
        {
            gate.Set();
        }

        Assert.True(factoryReturned.Wait(MaxElapsed));
    }

    [Fact]
    public void Track_InitializationTimeoutAlreadyElapsedSinceStart_DropsImmediately()
    {
        using var gate = new ManualResetEventSlim();
        using var factoryReturned = new ManualResetEventSlim();
        var sink = CreateSink(
            () =>
            {
                gate.Wait(BlockedFakeLimit);
                factoryReturned.Set();
                return new FakeSink();
            },
            initializationTimeout: SharedBudget);
        var items = NewItems();
        sink.Start();
        Thread.Sleep(SharedBudget + TimeSpan.FromMilliseconds(200));

        try
        {
            var elapsed = Measure(() => sink.Track(items.Event, items.Request));

            // The budget ran out while the command was working, so there is nothing left to wait for.
            Assert.True(elapsed < TimeSpan.FromMilliseconds(300), $"Track took {elapsed}.");
        }
        finally
        {
            gate.Set();
        }

        Assert.True(factoryReturned.Wait(MaxElapsed));
    }

    [Fact]
    public void Track_FactoryReturnsNull_IsANoOp()
    {
        var factoryCalls = 0;
        var sink = CreateSink(() =>
        {
            Interlocked.Increment(ref factoryCalls);
            return null;
        });
        var items = NewItems();

        var trackException = Record.Exception(() =>
        {
            sink.Track(items.Event, items.Request);
            sink.Track(items.Event, items.Request);
        });
        var shutdownException = Record.Exception(sink.Shutdown);

        Assert.Null(trackException);
        Assert.Null(shutdownException);
        Assert.Equal(1, factoryCalls);
    }

    [Fact]
    public void Track_FactoryThrows_IsANoOp()
    {
        var factoryCalls = 0;
        var sink = CreateSink(() =>
        {
            Interlocked.Increment(ref factoryCalls);
            throw new InvalidOperationException("SDK initialization failed");
        });
        var items = NewItems();

        var trackException = Record.Exception(() =>
        {
            sink.Track(items.Event, items.Request);
            sink.Track(items.Event, items.Request);
        });
        var shutdownException = Record.Exception(sink.Shutdown);

        Assert.Null(trackException);
        Assert.Null(shutdownException);
        Assert.Equal(1, factoryCalls);
    }

    [Fact]
    public void Track_SinkThrows_DoesNotThrow()
    {
        var fake = new FakeSink
        {
            OnTrack = (_, _) => throw new InvalidOperationException("track failed")
        };
        var sink = CreateSink(() => fake);
        var items = NewItems();

        var exception = Record.Exception(() => sink.Track(items.Event, items.Request));

        Assert.Null(exception);
    }

    private static DeferredTelemetrySink CreateSink(
        Func<ICliTelemetrySink?> factory,
        TimeSpan? initializationTimeout = null,
        TimeSpan? shutdownTimeout = null) =>
        new(factory, initializationTimeout ?? BlockedFakeLimit, shutdownTimeout ?? BlockedFakeLimit);

    private static (EventTelemetry Event, RequestTelemetry Request) NewItems(string name = "sheet/list") =>
        (new EventTelemetry(name), new RequestTelemetry { Name = name });

    private static TimeSpan Measure(Action action)
    {
        var stopwatch = Stopwatch.StartNew();
        action();
        return stopwatch.Elapsed;
    }

    private sealed class FakeSink : ICliTelemetrySink
    {
        private readonly object _gate = new();
        private readonly List<string> _calls = [];
        private readonly List<(EventTelemetry Event, RequestTelemetry Request)> _tracked = [];

        public Func<FakeSink, Task> FlushBehavior { get; init; } = CompleteFlushAsync;

        public Action OnDispose { get; init; } = () => { };

        public Action<EventTelemetry, RequestTelemetry>? OnTrack { get; init; }

        public IReadOnlyList<string> Calls
        {
            get
            {
                lock (_gate)
                {
                    return [.. _calls];
                }
            }
        }

        public IReadOnlyList<(EventTelemetry Event, RequestTelemetry Request)> Tracked
        {
            get
            {
                lock (_gate)
                {
                    return [.. _tracked];
                }
            }
        }

        public void Track(EventTelemetry eventTelemetry, RequestTelemetry requestTelemetry)
        {
            lock (_gate)
            {
                _calls.Add("track");
                _tracked.Add((eventTelemetry, requestTelemetry));
            }

            OnTrack?.Invoke(eventTelemetry, requestTelemetry);
        }

        public Task FlushAsync()
        {
            Record("flush");
            return FlushBehavior(this);
        }

        public void Dispose()
        {
            Record("dispose");
            OnDispose();
        }

        private void Record(string call)
        {
            lock (_gate)
            {
                _calls.Add(call);
            }
        }

        private static async Task CompleteFlushAsync(FakeSink sink)
        {
            await Task.Delay(20);
            sink.Record("flush-completed");
        }
    }
}
