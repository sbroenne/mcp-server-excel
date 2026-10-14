using System.Diagnostics;
using OpenTelemetry.Logs;
using OpenTelemetry.Trace;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Telemetry;
using Sbroenne.ExcelMcp.Telemetry;
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
public sealed class EntryPointTelemetryTests : IDisposable
{
    private readonly string? _previousMode = Environment.GetEnvironmentVariable("EXCELMCP_TELEMETRY_MODE");
    public EntryPointTelemetryTests() => Environment.SetEnvironmentVariable("EXCELMCP_TELEMETRY_MODE", "local");
    public void Dispose() => Environment.SetEnvironmentVariable("EXCELMCP_TELEMETRY_MODE", _previousMode);

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
    public async Task RunEntryPointAsync_UntrackedRoute_NeverStartsReporting(string commandLine)
    {
        var identityReads = 0;
        var reporter = new UsageTelemetry("ExcelMcp.Test", () =>
        {
            Interlocked.Increment(ref identityReads);
            return TelemetryUserIdentityStore.CreateProcessOnly();
        });
        using var scope = CliTelemetry.UseReporterForTesting(reporter);
        var daemonRuns = 0;
        var args = commandLine.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var exitCode = await Program.RunEntryPointAsync(args, CreateRuntime(_ => { daemonRuns++; return 0; }));
        await reporter.Initialization;
        Assert.Equal(0, exitCode);
        Assert.Equal(0, identityReads);
        Assert.False(reporter.IsReady);
        Assert.Equal(commandLine.StartsWith("service run", StringComparison.Ordinal) ? 1 : 0, daemonRuns);
    }

    [Fact]
    public async Task RunEntryPointAsync_ReadyReporter_RecordsFailureAndSignalsExitWithoutDisposingProviders()
    {
        var logs = new List<LogRecord>();
        var activities = new List<Activity>();
        var identity = TelemetryUserIdentityStore.CreateProcessOnly();
        var settings = new TelemetrySettings(TelemetryMode.Local, null, new Uri("http://localhost:4317"), "grpc");
        using var pipeline = new UsagePipeline(settings, "ExcelMcp.Test", identity, "test",
            t => t.AddInMemoryExporter(activities), l => l.AddInMemoryExporter(logs));
        var reporter = new UsageTelemetry("ExcelMcp.Test", () => identity, (_, _) => pipeline);
        using var scope = CliTelemetry.UseReporterForTesting(reporter);
        reporter.Start(settings);
        await reporter.Initialization;
        var exitCode = await Program.RunEntryPointAsync(
            ["--quiet", "sheet", "list", "--session", "entry-point-test"], CreateRuntime(_ => 0));
        Assert.NotEqual(0, exitCode);
        Assert.False(reporter.IsReady);
        Assert.True(pipeline.ForceFlush(1000));
        var request = Assert.Single(activities);
        Assert.Equal("sheet/list", request.DisplayName);
        Assert.Equal(ActivityStatusCode.Error, request.Status);
        Assert.Contains(logs, l => l.Body == "sheet/list");
    }

    [Fact]
    public async Task RunEntryPointAsync_BlockedReportingSetup_DoesNotWaitForReporting()
    {
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var reporter = new UsageTelemetry("ExcelMcp.Test", () =>
        {
            entered.Set();
            release.Wait();
            return TelemetryUserIdentityStore.CreateProcessOnly();
        });
        using var scope = CliTelemetry.UseReporterForTesting(reporter);
        var command = Task.Run(() => Program.RunEntryPointAsync(
            ["--quiet", "sheet", "list", "--session", "entry-point-test"], CreateRuntime(_ => 0)));
        try
        {
            Assert.True(entered.Wait(TimeSpan.FromSeconds(10)));
            Assert.Same(command, await Task.WhenAny(command, Task.Delay(500)));
            Assert.NotEqual(0, await command);
        }
        finally { release.Set(); await command; await reporter.Initialization; }
        Assert.False(reporter.IsReady);
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

}
