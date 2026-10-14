using System.Diagnostics;
using Sbroenne.ExcelMcp.Telemetry;
using Azure.Monitor.OpenTelemetry.Exporter;
using System.Text.Json;
using OpenTelemetry;
using OpenTelemetry.Logs;
using OpenTelemetry.Trace;
using Xunit;
#if CLI_TELEMETRY_TESTS
using Sbroenne.ExcelMcp.CLI.Telemetry;
#else
using Sbroenne.ExcelMcp.McpServer.Telemetry;
#endif

namespace Sbroenne.ExcelMcp.Tests.Telemetry;

[Trait("Category", "Unit")]
[Trait("Feature", "Telemetry")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
#if CLI_TELEMETRY_TESTS
[Collection("Sequential")]
[Trait("Layer", "CLI")]
#else
[Collection("ProgramTransport")]
[Trait("Layer", "McpServer")]
#endif
public sealed class LocalTelemetryTests
{
    private const string EmbeddedAzure = "InstrumentationKey=00000000-0000-0000-0000-000000000000;IngestionEndpoint=https://localhost.invalid/";

    [Theory]
    [InlineData(null, null, "Off")]
    [InlineData("__placeholder__", null, "Off")]
    [InlineData(EmbeddedAzure, null, "Azure")]
    [InlineData(EmbeddedAzure, "off", "Off")]
    [InlineData(null, "local", "Local")]
    [InlineData(EmbeddedAzure, "local", "Local")]
    [InlineData(EmbeddedAzure, "azure", "Azure")]
    [InlineData(EmbeddedAzure, " LOCAL ", "Local")]
    public void Resolve_Mode_SelectsOnlyRequestedDestination(string? embedded, string? mode, string expected)
    {
        var settings = TelemetrySettings.Resolve(embedded, key => key == TelemetrySettings.ModeVariable ? mode : null);
        Assert.Equal(expected, settings.Mode.ToString());
        Assert.Equal(expected == "Azure" ? EmbeddedAzure : null, settings.ConnectionString);
        Assert.Equal(expected == "Local" ? new Uri("http://localhost:4317") : null, settings.Endpoint);
    }

    [Theory]
    [InlineData("unknown", null, null)]
    [InlineData("azure", null, null)]
    [InlineData("local", "https://example.com", "grpc")]
    [InlineData("local", "http://localhost:4317?secret=value", "grpc")]
    [InlineData("local", "http://user:password@localhost:4317", "grpc")]
    [InlineData("local", "not-a-uri", "grpc")]
    [InlineData("local", "http://localhost:4317", "http/json")]
    public void Resolve_InvalidConfiguration_ReportsSafeError(string mode, string? endpoint, string? protocol)
    {
        var variables = new Dictionary<string, string?>
        {
            [TelemetrySettings.ModeVariable] = mode,
            ["OTEL_EXPORTER_OTLP_ENDPOINT"] = endpoint,
            ["OTEL_EXPORTER_OTLP_PROTOCOL"] = protocol
        };
        var error = Assert.Throws<InvalidOperationException>(() =>
            TelemetrySettings.Resolve(null, key => variables.GetValueOrDefault(key)));
        Assert.DoesNotContain("secret=value", error.Message);
        Assert.DoesNotContain("password", error.Message);
        Assert.NotEmpty(error.Message);
    }


    private static readonly TelemetryUserIdentity Identity = new("0123456789abcdef0123456789abcdef", "abcdef0123456789");
    private static TelemetrySettings Local => new(TelemetryMode.Local, null, new Uri("http://localhost:4317"), "grpc");

    [Fact]
    public async Task Start_BlockedIdentity_DoesNotBlockRecordsOrExit_AndCannotBecomeReadyAfterExit()
    {
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var builds = 0;
        var reporter = new UsageTelemetry("ExcelMcp.Test", () =>
        {
            entered.Set();
            release.Wait();
            return Identity;
        }, (_, _) => { Interlocked.Increment(ref builds); throw new InvalidOperationException(); });
        var call = Task.Run(() => reporter.Start(Local));
        try
        {
            Assert.True(entered.Wait(TimeSpan.FromSeconds(10)));
            await call.WaitAsync(TimeSpan.FromMilliseconds(500));
            var recordAndExit = Task.Run(() =>
            {
                reporter.RecordEvent(new UsageEvent("before-ready"));
                reporter.RecordOperation(new UsageOperation { Name = "before-ready" });
                reporter.SignalExit();
            });
            await recordAndExit.WaitAsync(TimeSpan.FromMilliseconds(500));
            Assert.False(reporter.IsReady);
        }
        finally { release.Set(); await call; await reporter.Initialization; }
        Assert.Equal(0, builds);
        Assert.False(reporter.IsReady);
    }

    [Fact]
    public async Task Start_Disabled_NeverAccessesIdentityOrBuildsProvider()
    {
        var identityCalls = 0;
        var providerCalls = 0;
        var reporter = new UsageTelemetry("ExcelMcp.Test", () =>
        {
            Interlocked.Increment(ref identityCalls);
            return Identity;
        }, (_, _) =>
        {
            Interlocked.Increment(ref providerCalls);
            throw new InvalidOperationException("Provider was built");
        });
        reporter.Start(new(TelemetryMode.Off, null, null, null));
        await reporter.Initialization;
        reporter.RecordEvent(new UsageEvent("disabled"));
        Assert.False(reporter.IsReady);
        Assert.Same(Task.CompletedTask, reporter.Initialization);
        Assert.Equal(0, identityCalls);
        Assert.Equal(0, providerCalls);
    }

    [Fact]
    public async Task Ready_RecordsIdentityMigrationExactlyOnce_AndDropsBeforeReadyAndAfterExit()
    {
        var activities = new List<Activity>();
        var logs = new List<LogRecord>();
        using var pipeline = new UsagePipeline(Local, "ExcelMcp.Test", Identity, "test-session",
            t => t.AddInMemoryExporter(activities), l => l.AddInMemoryExporter(logs));
        var reporter = new UsageTelemetry("ExcelMcp.Test", () => Identity, (_, _) => pipeline);
        reporter.RecordEvent(new UsageEvent("before-ready"));
        reporter.Start(Local);
        reporter.Start(Local);
        await reporter.Initialization;
        Assert.True(reporter.IsReady);
        var operation = new UsageOperation
        {
            Name = "sheet/list",
            Success = true,
            Timestamp = DateTimeOffset.UtcNow.AddMilliseconds(-123),
            Duration = TimeSpan.FromMilliseconds(123)
        };
        operation.Properties["Outcome"] = "expected-negative";
        reporter.RecordOperation(operation);
        reporter.SignalExit();
        reporter.RecordEvent(new UsageEvent("after-exit"));
        reporter.RecordOperation(new UsageOperation { Name = "after-exit" });
        Assert.True(pipeline.ForceFlush(1000));
        var activity = Assert.Single(activities);
        Assert.Equal(operation.Timestamp.UtcDateTime, activity.StartTimeUtc);
        Assert.Equal(operation.Duration, activity.Duration);
        Assert.Equal(ActivityStatusCode.Ok, activity.Status);
        Assert.Equal(Identity.UserId, activity.GetTagItem("UserId"));
        Assert.Equal("test-session", activity.GetTagItem("SessionId"));
        Assert.Equal("expected-negative", activity.GetTagItem("Outcome"));
        Assert.Equal(2, logs.Count);
        var migration = Assert.Single(logs, l => l.Body == "UserIdMigration");
        Assert.Contains(migration.Attributes!, a => a.Key == "LegacyUserId" && Equals(a.Value, Identity.LegacyUserId));
        Assert.Contains(logs, l => l.Body == "SessionStart");
    }

    [Fact]
    public async Task BatchProcessors_BlockedExportAndFullBuffers_NeverBlockRecordingOrSignalExit()
    {
        using var release = new ManualResetEventSlim();
        using var traceExporter = new BlockingExporter<Activity>(release);
        using var logExporter = new BlockingExporter<LogRecord>(release);
        using var pipeline = new UsagePipeline(Local, "ExcelMcp.Test", Identity, "test-session",
            t => t.AddProcessor(new BatchActivityExportProcessor(traceExporter, maxQueueSize: 16,
                scheduledDelayMilliseconds: 1, maxExportBatchSize: 1)),
            l => l.AddProcessor(new BatchLogRecordExportProcessor(logExporter, maxQueueSize: 16,
                scheduledDelayMilliseconds: 1, maxExportBatchSize: 1)));
        var reporter = new UsageTelemetry("ExcelMcp.Test", () => Identity, (_, _) => pipeline);
        reporter.Start(Local);
        await reporter.Initialization;
        reporter.RecordOperation(new UsageOperation { Name = "sheet/list", Success = true });
        try
        {
            Assert.True(traceExporter.Entered.Wait(TimeSpan.FromSeconds(10)));
            Assert.True(logExporter.Entered.Wait(TimeSpan.FromSeconds(10)));
            var calls = Task.Run(() =>
            {
                for (var i = 0; i < 4096; i++)
                {
                    reporter.RecordEvent(new UsageEvent("sheet/list"));
                    reporter.RecordOperation(new UsageOperation { Name = "sheet/list", Success = true });
                }
                reporter.SignalExit();
            });
            await calls.WaitAsync(TimeSpan.FromSeconds(1));
            Assert.False(reporter.IsReady);
        }
        finally { release.Set(); }
        Assert.True(pipeline.ForceFlush(10000));
        Assert.InRange(traceExporter.Count, 1L, 17L);
        Assert.InRange(logExporter.Count, 1L, 17L);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task LocalExport_DeliversExplicitSafeFields_WithoutAmbientResourceOrHeaders(bool grpc)
    {
        await using var http = new LocalOtlpReceiver();
        await using var grpcReceiver = await LocalGrpcOtlpReceiver.StartAsync();
        var previous = Environment.GetEnvironmentVariable("OTEL_RESOURCE_ATTRIBUTES");
        var headerVariables = new[] { "OTEL_EXPORTER_OTLP_HEADERS", "OTEL_EXPORTER_OTLP_LOGS_HEADERS", "OTEL_EXPORTER_OTLP_TRACES_HEADERS" };
        var previousHeaders = headerVariables.ToDictionary(name => name, Environment.GetEnvironmentVariable);
        try
        {
            Environment.SetEnvironmentVariable("OTEL_RESOURCE_ATTRIBUTES", "private_path=confidential-workbook.xlsx");
            foreach (var variable in headerVariables) Environment.SetEnvironmentVariable(variable, "x-private=secret");
            using var pipeline = new UsagePipeline(new(TelemetryMode.Local, null,
                new Uri(grpc ? grpcReceiver.Endpoint : http.Endpoint), grpc ? "grpc" : "http/protobuf"),
                "ExcelMcp.Test", Identity, "test-session");
            pipeline.RecordEvent(new UsageEvent("sheet/list"));
            pipeline.RecordOperation(new UsageOperation
            {
                Name = "sheet/list",
                Success = false,
                Timestamp = DateTimeOffset.UtcNow.AddMilliseconds(-15),
                Duration = TimeSpan.FromMilliseconds(15)
            });
            Assert.True(pipeline.ForceFlush(10000));
            var requests = grpc ? grpcReceiver.Requests : http.Requests;
            Assert.Equal(2, requests.Count);
            Assert.All(requests, request =>
            {
                Assert.Contains("sheet/list", request.Payload);
                Assert.Contains(Identity.UserId, request.Payload);
                Assert.DoesNotContain("confidential-workbook.xlsx", request.Payload);
                Assert.DoesNotContain("x-private", request.Headers.Keys, StringComparer.OrdinalIgnoreCase);
            });
        }
        finally
        {
            Environment.SetEnvironmentVariable("OTEL_RESOURCE_ATTRIBUTES", previous);
            foreach (var variable in previousHeaders) Environment.SetEnvironmentVariable(variable.Key, variable.Value);
        }
    }

    [Fact]
    public async Task LocalExport_Redirect_DoesNotForwardTelemetry()
    {
        await using var destination = new LocalOtlpReceiver();
        await using var redirect = new LocalOtlpReceiver(redirect: destination.Endpoint);
        using var pipeline = new UsagePipeline(new(TelemetryMode.Local, null,
            new Uri(redirect.Endpoint), "http/protobuf"), "ExcelMcp.Test", Identity, "test-session");
        pipeline.RecordEvent(new UsageEvent("sheet/list"));
        Assert.True(pipeline.ForceFlush(10000));
        Assert.Single(redirect.Requests);
        Assert.Empty(destination.Requests);
    }

    [Fact]
    public async Task AzureExporter_BurstIsNotRateLimitedOrSampled()
    {
        await using var receiver = new LocalOtlpReceiver(responseBody:
            """{"itemsReceived":100,"itemsAccepted":100,"errors":[]}""");
        using var pipeline = new UsagePipeline(new(TelemetryMode.Azure,
            $"InstrumentationKey=00000000-0000-0000-0000-000000000000;IngestionEndpoint={receiver.Endpoint}/", null, null),
            "ExcelMcp.Test", Identity, "test-session");
        for (var i = 0; i < 100; i++)
        {
            pipeline.RecordEvent(new UsageEvent("sheet/list"));
            pipeline.RecordOperation(new UsageOperation
            {
                Name = "sheet/list",
                Success = true,
                Timestamp = DateTimeOffset.UtcNow.AddMilliseconds(-1),
                Duration = TimeSpan.FromMilliseconds(1)
            });
        }
        Assert.True(pipeline.ForceFlush(10000));
        var items = receiver.Requests.SelectMany(r => r.Payload.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            .Select(s => JsonSerializer.Deserialize<JsonElement>(s)).ToArray();
        Assert.Equal(100, items.Count(i => i.GetProperty("data").GetProperty("baseType").GetString() == "RequestData"));
        Assert.Equal(100, items.Count(i => i.GetProperty("data").GetProperty("baseType").GetString() == "EventData"));
    }

    [Fact]
    public void AzureOptions_DisableSamplingDiskMetricsAndSelfStatistics()
    {
        var options = new AzureMonitorExporterOptions();
        UsagePipeline.ConfigureAzure(options, new(TelemetryMode.Azure, EmbeddedAzure, null, null));
        Assert.Equal(1f, options.SamplingRatio);
        Assert.Null(options.TracesPerSecond);
        Assert.True(options.DisableOfflineStorage);
        Assert.False(options.EnableLiveMetrics);
        Assert.False(options.EnableStandardMetrics);
        Assert.False(options.EnablePerformanceCounters);
        Assert.Equal("true", Environment.GetEnvironmentVariable("APPLICATIONINSIGHTS_SDKSTATS_DISABLED"));
        Assert.Equal("true", Environment.GetEnvironmentVariable("APPLICATIONINSIGHTS_STATSBEAT_DISABLED"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task AzureExporter_UsesNativeEventsRequestsAndAnonymousResource_WithoutAutomaticMetrics(bool succeeded)
    {
        await using var receiver = new LocalOtlpReceiver(responseBody:
            """{"itemsReceived":2,"itemsAccepted":2,"errors":[]}""");
        using var pipeline = new UsagePipeline(new(TelemetryMode.Azure,
            $"InstrumentationKey=00000000-0000-0000-0000-000000000000;IngestionEndpoint={receiver.Endpoint}/", null, null),
            "ExcelMcp.Test", Identity, "test-session");
        pipeline.RecordEvent(new UsageEvent("sheet/list"));
        pipeline.RecordOperation(new UsageOperation
        {
            Name = "sheet/list",
            Success = succeeded,
            Timestamp = DateTimeOffset.UtcNow.AddMilliseconds(-123),
            Duration = TimeSpan.FromMilliseconds(123)
        });
        Assert.True(pipeline.ForceFlush(10000));
        var items = receiver.Requests.SelectMany(r => r.Payload.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            .Select(s => JsonSerializer.Deserialize<JsonElement>(s)).ToArray();
        // Azure includes its resource envelope, not standard metrics or counters.
        Assert.Equal(3, items.Length);
        var resource = Assert.Single(items, i => i.GetProperty("data").GetProperty("baseType").GetString() == "MetricData");
        Assert.Equal("_OTELRESOURCE_", Assert.Single(resource.GetProperty("data").GetProperty("baseData")
            .GetProperty("metrics").EnumerateArray()).GetProperty("name").GetString());
        var eventItem = Assert.Single(items, i => i.GetProperty("data").GetProperty("baseType").GetString() == "EventData");
        var eventData = eventItem.GetProperty("data").GetProperty("baseData");
        Assert.Equal("sheet/list", eventData.GetProperty("name").GetString());
        Assert.Equal(Identity.UserId, eventData.GetProperty("properties").GetProperty("UserId").GetString());
        Assert.Equal("test-session", eventData.GetProperty("properties").GetProperty("SessionId").GetString());
        var request = Assert.Single(items, i => i.GetProperty("data").GetProperty("baseType").GetString() == "RequestData");
        var data = request.GetProperty("data").GetProperty("baseData");
        Assert.Equal("sheet/list", data.GetProperty("name").GetString());
        Assert.Equal("0", data.GetProperty("responseCode").GetString());
        Assert.Equal(succeeded, data.GetProperty("success").GetBoolean());
        Assert.Equal(Identity.UserId, data.GetProperty("properties").GetProperty("UserId").GetString());
        Assert.Equal("instance-01234567", request.GetProperty("tags").GetProperty("ai.cloud.roleInstance").GetString());
        Assert.All(receiver.Requests, r => Assert.DoesNotContain(Environment.MachineName, r.Payload, StringComparison.OrdinalIgnoreCase));
    }

    private sealed class BlockingExporter<T>(ManualResetEventSlim release) : BaseExporter<T> where T : class
    {
        internal ManualResetEventSlim Entered { get; } = new();
        internal long Count { get; private set; }
        public override ExportResult Export(in Batch<T> batch)
        {
            Entered.Set();
            release.Wait();
            Count += batch.Count;
            return ExportResult.Success;
        }
    }
}
