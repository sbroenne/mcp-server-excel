using Microsoft.ApplicationInsights.DataContracts;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Telemetry;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

/// <summary>
/// Runs the same path as <c>Program.Main</c> so a regression that starts telemetry
/// for untracked routes (help, version, no command, the background service) fails here.
/// </summary>
// Replaces the process-wide telemetry sink, so it must not overlap other collections.
[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "Telemetry")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class EntryPointTelemetryTests
{
    [Theory]
    [InlineData("--help")]
    [InlineData("-h")]
    [InlineData("sheet --help")]
    [InlineData("--version")]
    [InlineData("-v")]
    [InlineData("")]
    [InlineData("--quiet")]
    [InlineData("service run")]
    [InlineData("service run --pipe-name entry-point-test")]
    public async Task RunEntryPointAsync_UntrackedRoute_NeverCreatesTelemetrySink(string commandLine)
    {
        var args = commandLine.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var factoryCalls = 0;
        var fake = new RecordingSink();
        string? daemonPipeName = "not-run";
        var daemonRuns = 0;
        using var telemetryScope = CliTelemetry.UseSinkFactoryForTesting(() =>
        {
            Interlocked.Increment(ref factoryCalls);
            return fake;
        });

        var exitCode = await Program.RunEntryPointAsync(
            args,
            CreateRuntime(pipeName =>
            {
                daemonRuns++;
                daemonPipeName = pipeName;
                return 0;
            }));

        Assert.Equal(0, exitCode);
        Assert.Equal(0, factoryCalls);
        Assert.Empty(fake.Tracked);
        if (args.Length >= 2 && args[0] == "service" && args[1] == "run")
        {
            // Proves the service route was actually taken, not short-circuited elsewhere.
            Assert.Equal(1, daemonRuns);
            Assert.Equal(args.Length == 4 ? args[3] : null, daemonPipeName);
        }
        else
        {
            Assert.Equal(0, daemonRuns);
        }
    }

    [Fact]
    public async Task RunEntryPointAsync_TrackedCommand_CreatesSinkOnceAndSendsInvocation()
    {
        var factoryCalls = 0;
        var fake = new RecordingSink();
        using var telemetryScope = CliTelemetry.UseSinkFactoryForTesting(() =>
        {
            Interlocked.Increment(ref factoryCalls);
            return fake;
        });

        // The stub service connection fails, so the command is tracked as a failure
        // without needing a running daemon.
        var exitCode = await Program.RunEntryPointAsync(
            ["--quiet", "sheet", "list", "--session", "entry-point-test"],
            CreateRuntime(_ => throw new InvalidOperationException("Daemon must not run.")));

        Assert.NotEqual(0, exitCode);
        Assert.Equal(1, factoryCalls);
        var tracked = Assert.Single(fake.Tracked);
        Assert.Equal("sheet/list", tracked.Event.Name);
        Assert.Equal("sheet/list", tracked.Request.Name);
        Assert.False(tracked.Request.Success);
        // RunEntryPointAsync flushes on exit, as Main does.
        Assert.Equal(["flush", "dispose"], fake.Lifecycle);
    }

    private static CliCommandRuntime CreateRuntime(Func<string?, int> serviceDaemonRunner) =>
        new(
            new FailingClientFactory(),
            new StringReader(string.Empty),
            new StringWriter(),
            new StringWriter(),
            isOutputRedirected: true,
            latestVersionProvider: () => Task.FromResult<string?>(null),
            serviceDaemonRunner: serviceDaemonRunner);

    private sealed class FailingClientFactory : ICliRequestClientFactory
    {
        public Task<ICliRequestClient> ConnectAsync(CancellationToken cancellationToken) =>
            throw new InvalidOperationException("No service is available in this test.");
    }

    private sealed class RecordingSink : ICliTelemetrySink
    {
        private readonly object _gate = new();
        private readonly List<(EventTelemetry Event, RequestTelemetry Request)> _tracked = [];
        private readonly List<string> _lifecycle = [];

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

        public IReadOnlyList<string> Lifecycle
        {
            get
            {
                lock (_gate)
                {
                    return [.. _lifecycle];
                }
            }
        }

        public void Track(EventTelemetry eventTelemetry, RequestTelemetry requestTelemetry)
        {
            lock (_gate)
            {
                _tracked.Add((eventTelemetry, requestTelemetry));
            }
        }

        public Task FlushAsync()
        {
            Record("flush");
            return Task.CompletedTask;
        }

        public void Dispose() => Record("dispose");

        private void Record(string call)
        {
            lock (_gate)
            {
                _lifecycle.Add(call);
            }
        }
    }
}
