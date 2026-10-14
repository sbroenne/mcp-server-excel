using OpenTelemetry.Logs;
using OpenTelemetry.Trace;
using System.Diagnostics;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Telemetry;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Collection("ProgramTransport")]
[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "Telemetry")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class LocalTelemetryConfigurationTests
{
    [Fact]
    public async Task RunAsync_BlockedIdentity_DoesNotDelayStartupRequestsOrShutdown()
    {
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var reporter = new UsageTelemetry("ExcelMcp.Test", () =>
        {
            entered.Set();
            release.Wait();
            return TelemetryUserIdentityStore.CreateProcessOnly();
        });
        using var scope = ExcelMcpTelemetry.UseReporterForTesting(reporter);
        var previous = Environment.GetEnvironmentVariable("EXCELMCP_TELEMETRY_MODE");
        var fixture = new RecordingProgramTransportFixture();
        try
        {
            Environment.SetEnvironmentVariable("EXCELMCP_TELEMETRY_MODE", "local");
            await fixture.InitializeAsync().WaitAsync(TimeSpan.FromSeconds(10));
            Assert.True(entered.Wait(TimeSpan.FromSeconds(10)));
            await fixture.CallToolAsync("file_read", new Dictionary<string, object?> { ["action"] = "list" },
                new ServiceResponse { Success = true, Result = """{"success":true,"sessions":[]}""" },
                "session.list", expectedSessionId: null, expectedArgsJson: null).WaitAsync(TimeSpan.FromSeconds(2));
            await fixture.DisposeAsync().WaitAsync(TimeSpan.FromSeconds(2));
            Assert.False(reporter.IsReady);
        }
        finally
        {
            release.Set();
            await reporter.Initialization;
            await fixture.DisposeAsync();
            Environment.SetEnvironmentVariable("EXCELMCP_TELEMETRY_MODE", previous);
        }
    }

    [Fact]
    public async Task RunAsync_ReadyReporter_RecordsSafeUsageAndException_WithoutHostDisposal()
    {
        var activities = new List<Activity>();
        var logs = new List<LogRecord>();
        var identity = TelemetryUserIdentityStore.CreateProcessOnly();
        var settings = new TelemetrySettings(TelemetryMode.Local, null, new Uri("http://localhost:4317"), "grpc");
        using var pipeline = new UsagePipeline(settings, "ExcelMcp.Test", identity, "test-session",
            t => t.AddInMemoryExporter(activities), l => l.AddInMemoryExporter(logs));
        var reporter = new UsageTelemetry("ExcelMcp.Test", () => identity, (_, _) => pipeline);
        using var scope = ExcelMcpTelemetry.UseReporterForTesting(reporter);
        reporter.Start(settings);
        await reporter.Initialization;
        await using (var fixture = new RecordingProgramTransportFixture())
        {
            await fixture.InitializeAsync();
            await fixture.CallToolAsync("file_read", new Dictionary<string, object?> { ["action"] = "list" },
                new ServiceResponse { Success = true, Result = """{"success":true,"sessions":[]}""" },
                "session.list", expectedSessionId: null, expectedArgsJson: null);
            ExcelMcpTelemetry.TrackUnhandledException(new InvalidOperationException("confidential-workbook.xlsx customer@example.com"), "McpServer.RunAsync");
        }
        Assert.False(reporter.IsReady);
        Assert.True(pipeline.ForceFlush(1000));
        Assert.Equal("file_read/list", Assert.Single(activities).DisplayName);
        var error = Assert.Single(logs, l => l.Body == "UnhandledException");
        Assert.Contains(error.Attributes!, a => a.Key == "Sanitized" && Equals(a.Value, "true"));
        Assert.Single(error.Attributes!, a => a.Key == "AppVersion");
        Assert.DoesNotContain("confidential-workbook.xlsx", string.Join("\n", error.Attributes!));
        Assert.DoesNotContain("customer@example.com", string.Join("\n", error.Attributes!));
    }
}
