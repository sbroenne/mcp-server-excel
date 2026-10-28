using System.Diagnostics;
using System.Text.Json;
using OpenTelemetry.Logs;
using OpenTelemetry.Trace;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Telemetry;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Telemetry;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

/// <summary>
/// Runs the same path as <c>Program.Main</c> so a regression that starts telemetry
/// for untracked routes (help, version, no command) fails here.
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
    [InlineData("sheet list --session entry-point-test")]
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
        if (commandLine.StartsWith("sheet", StringComparison.Ordinal) &&
            !commandLine.Contains("--help", StringComparison.Ordinal))
            Assert.NotEqual(0, exitCode);
        else
            Assert.Equal(0, exitCode);
        Assert.Equal(0, identityReads);
        Assert.False(reporter.IsReady);
        Assert.Equal(0, daemonRuns);
    }

    [Theory]
    [InlineData("--fields running", "--fields")]
    [InlineData("--fields=running", "--fields=running")]
    [InlineData("--pipe-name test-pipe --fields running", "--fields")]
    [InlineData("--pipe-name", "--pipe-name")]
    [InlineData("--pipe-name --fields", "--pipe-name")]
    [InlineData("--pipe-name -h", "--pipe-name")]
    [InlineData("--unknown", "--unknown")]
    public async Task RunEntryPointAsync_InvalidDaemonArguments_DoesNotStartHostOrReporting(
        string arguments, string invalidArgument)
    {
        var identityReads = 0;
        var reporter = new UsageTelemetry("ExcelMcp.Test", () =>
        {
            Interlocked.Increment(ref identityReads);
            return TelemetryUserIdentityStore.CreateProcessOnly();
        });
        using var scope = CliTelemetry.UseReporterForTesting(reporter);
        var daemonRuns = 0;
        var runtime = CreateRuntime(_ => { daemonRuns++; return 0; });

        var exitCode = await Program.RunEntryPointAsync(
            ["service", "run", .. arguments.Split(' ')], runtime);
        await reporter.Initialization;

        Assert.Equal(0, daemonRuns);
        Assert.NotEqual(0, exitCode);
        Assert.Equal(0, identityReads);
        using var output = JsonDocument.Parse(Assert.IsType<StringWriter>(runtime.Output).ToString());
        Assert.False(output.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("InvalidInput", output.RootElement.GetProperty("error").GetProperty("category").GetString());
        Assert.Contains(invalidArgument, output.RootElement.GetProperty("error").GetProperty("message").GetString());
    }

    [Theory]
    [InlineData("", null)]
    [InlineData("--pipe-name test-pipe", "test-pipe")]
    [InlineData("--PIPE-NAME test-pipe", "test-pipe")]
    [InlineData("--pipe-name -", "-")]
    public async Task RunEntryPointAsync_ValidDaemonArguments_PreservePipeAndExitCode(
        string arguments, string? expectedPipeName)
    {
        var reporter = new UsageTelemetry("ExcelMcp.Test", TelemetryUserIdentityStore.CreateProcessOnly);
        using var scope = CliTelemetry.UseReporterForTesting(reporter);
        var daemonRuns = 0;
        var runtime = CreateRuntime(pipeName =>
        {
            daemonRuns++;
            Assert.Equal(expectedPipeName, pipeName);
            return 7;
        });

        var exitCode = await Program.RunEntryPointAsync(
            ["service", "run", .. arguments.Split(' ', StringSplitOptions.RemoveEmptyEntries)], runtime);
        await reporter.Initialization;

        Assert.Equal(1, daemonRuns);
        Assert.Equal(7, exitCode);
        Assert.Empty(Assert.IsType<StringWriter>(runtime.Output).ToString());
    }

    [Fact]
    public async Task DaemonPipe_RecordsEachOperationOnce_WithoutReportingInProcessCallsOrReadinessPings()
    {
        var logs = new List<LogRecord>();
        var activities = new List<Activity>();
        var identity = TelemetryUserIdentityStore.CreateProcessOnly();
        var settings = new TelemetrySettings(TelemetryMode.Local, null, new Uri("http://localhost:4317"), "grpc");
        using var pipeline = new UsagePipeline(settings, "ExcelMcp.Test", identity, "daemon-test",
            t => t.AddInMemoryExporter(activities), l => l.AddInMemoryExporter(logs));
        var reporter = new UsageTelemetry("ExcelMcp.Test", () => identity, (_, _) => pipeline);
        using var scope = CliTelemetry.UseReporterForTesting(reporter);
        reporter.Start(settings);
        await reporter.Initialization;
        using var service = new ExcelMcpService(CliTelemetry.TrackDaemonRequestAsync);
        Assert.True((await service.ProcessAsync(new ServiceRequest { Command = "session.list" })).Success);
        Assert.Empty(activities);

        var pipe = $"daemon-observer-test-{Guid.NewGuid():N}";
        var running = service.RunAsync(pipe);
        using var client = new ServiceClient(pipe);
        try
        {
            Assert.True((await client.SendAsync(new ServiceRequest { Command = "service.ping" })).Success);
            Assert.True((await client.SendAsync(new ServiceRequest { Command = "session.list" })).Success);
            var failed = await client.SendAsync(
                new ServiceRequest { Command = "sheet.list", SessionId = "missing-session" });
            Assert.False(failed.Success);
            Assert.Equal("SessionNotFound", failed.ErrorCategory);
            Assert.True(pipeline.ForceFlush(1000));
            Assert.Equal(2, activities.Count);
            Assert.Equal(ActivityStatusCode.Ok, Assert.Single(activities, a => a.DisplayName == "session/list").Status);
            Assert.Equal(ActivityStatusCode.Error, Assert.Single(activities, a => a.DisplayName == "sheet/list").Status);
            Assert.Single(logs, l => l.Body == "session/list");
            Assert.Single(logs, l => l.Body == "sheet/list");
            Assert.DoesNotContain(logs, l => l.Body == "service/ping");
        }
        finally
        {
            service.RequestShutdown();
            await running.WaitAsync(TimeSpan.FromSeconds(10));
        }
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
    public async Task RunEntryPointAsync_BlockedReportingSetup_DoesNotWaitForDaemonStartup()
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
            ["service", "run", "--pipe-name", "entry-point-test"], CreateRuntime(_ => 0)));
        try
        {
            Assert.True(entered.Wait(TimeSpan.FromSeconds(10)));
            Assert.Same(command, await Task.WhenAny(command, Task.Delay(500)));
            Assert.Equal(0, await command);
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
