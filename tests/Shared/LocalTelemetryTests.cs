using System.Diagnostics;
using Microsoft.ApplicationInsights.DataContracts;
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

    [Fact]
    public void LocalPipeline_RecordsEventAndRequestOnce_WithExactFieldsAndDuration()
    {
        var activities = new List<Activity>();
        var logs = new List<LogRecord>();
        var settings = TelemetrySettings.Resolve(EmbeddedAzure, key => key == TelemetrySettings.ModeVariable ? "local" : null);
        using var telemetry = new LocalTelemetry(settings, "ExcelMcp.Test",
            tracing => tracing.AddInMemoryExporter(activities),
            logging => logging.AddInMemoryExporter(logs));
        var eventTelemetry = new EventTelemetry("sheet/list");
        eventTelemetry.Properties["Outcome"] = "expected-negative";
        eventTelemetry.Properties["EntryPoint"] = "test";
        eventTelemetry.Context.User.Id = "anonymous-test-user";
        eventTelemetry.Context.Session.Id = "test-session";
        var request = new RequestTelemetry
        {
            Name = "sheet/list",
            Success = true,
            ResponseCode = "200",
            Timestamp = DateTimeOffset.UtcNow.AddMilliseconds(-123),
            Duration = TimeSpan.FromMilliseconds(123)
        };
        request.Properties["Outcome"] = "expected-negative";
        telemetry.Track(eventTelemetry);
        telemetry.Track(request);
        Assert.True(telemetry.Flush());

        var activity = Assert.Single(activities);
        Assert.Equal("sheet/list", activity.DisplayName);
        Assert.Equal(ActivityStatusCode.Ok, activity.Status);
        Assert.Equal(request.Timestamp.UtcDateTime, activity.StartTimeUtc);
        Assert.Equal(request.Duration, activity.Duration);
        Assert.Equal("expected-negative", activity.GetTagItem("Outcome"));
        Assert.Equal("200", activity.GetTagItem("ResponseCode"));
        var log = Assert.Single(logs);
        Assert.Equal("sheet/list", log.Body);
        Assert.Contains(log.Attributes!, attribute => attribute.Key == "UserId" && Equals(attribute.Value, "anonymous-test-user"));
        Assert.Contains(log.Attributes!, attribute => attribute.Key == "Outcome" && Equals(attribute.Value, "expected-negative"));
    }

    [Fact]
    public void LocalPipeline_RejectsUnsanitizedException()
    {
        using var telemetry = new LocalTelemetry(
            TelemetrySettings.Resolve(null, key => key == TelemetrySettings.ModeVariable ? "local" : null), "ExcelMcp.Test",
            tracing => tracing.AddInMemoryExporter(new List<Activity>()),
            logging => logging.AddInMemoryExporter(new List<LogRecord>()));
        Assert.Throws<InvalidOperationException>(() =>
            telemetry.Track(new ExceptionTelemetry(new InvalidOperationException("private workbook contents"))));
    }

    [Fact]
    public async Task LocalPipeline_HttpProtobuf_DeliversEventsAndRequestsToLoopback()
    {
        await using var receiver = new LocalOtlpReceiver();
        var variables = new Dictionary<string, string>
        {
            [TelemetrySettings.ModeVariable] = "local",
            ["OTEL_EXPORTER_OTLP_ENDPOINT"] = receiver.Endpoint,
            ["OTEL_EXPORTER_OTLP_PROTOCOL"] = "http/protobuf"
        };
        using var telemetry = new LocalTelemetry(
            TelemetrySettings.Resolve(EmbeddedAzure, key => variables.GetValueOrDefault(key)), "ExcelMcp.Test");
        var eventTelemetry = new EventTelemetry("worksheet_read/list");
        eventTelemetry.Properties["Outcome"] = "succeeded";
        telemetry.Track(eventTelemetry);
        telemetry.Track(new RequestTelemetry
        {
            Name = "worksheet_read/list",
            Timestamp = DateTimeOffset.UtcNow.AddMilliseconds(-10),
            Duration = TimeSpan.FromMilliseconds(10),
            Success = true,
            ResponseCode = "200"
        });
        Assert.True(telemetry.Flush());
        var requests = receiver.Requests;
        Assert.Equal(2, requests.Count);
        Assert.Contains(requests, request => request.Path == "/v1/traces" && request.Payload.Contains("worksheet_read/list", StringComparison.Ordinal));
        Assert.Contains(requests, request => request.Path == "/v1/logs" && request.Payload.Contains("Outcome", StringComparison.Ordinal));
        Assert.All(requests, request => Assert.Contains("ExcelMcp.Test", request.Payload));
        Assert.All(requests, request => Assert.DoesNotContain("localhost.invalid", request.Payload));
    }

    [Fact]
    public async Task LocalPipeline_DefaultGrpc_DeliversEventsAndRequestsToLoopback()
    {
        await using var receiver = await LocalGrpcOtlpReceiver.StartAsync();
        var variables = new Dictionary<string, string>
        {
            [TelemetrySettings.ModeVariable] = "local",
            ["OTEL_EXPORTER_OTLP_ENDPOINT"] = receiver.Endpoint
        };
        var settings = TelemetrySettings.Resolve(EmbeddedAzure, key => variables.GetValueOrDefault(key));
        Assert.Equal("grpc", settings.Protocol);
        using var telemetry = new LocalTelemetry(settings, "ExcelMcp.Test");
        var eventTelemetry = new EventTelemetry("worksheet_read/list");
        eventTelemetry.Properties["Outcome"] = "succeeded";
        telemetry.Track(eventTelemetry);
        telemetry.Track(new RequestTelemetry
        {
            Name = "worksheet_read/list",
            Timestamp = DateTimeOffset.UtcNow.AddMilliseconds(-10),
            Duration = TimeSpan.FromMilliseconds(10),
            Success = true,
            ResponseCode = "200"
        });
        Assert.True(telemetry.Flush());
        var requests = receiver.Requests;
        Assert.Equal(2, requests.Count);
        Assert.Contains(requests, request => request.Path == "/opentelemetry.proto.collector.trace.v1.TraceService/Export"
            && request.Payload.Contains("worksheet_read/list", StringComparison.Ordinal));
        Assert.Contains(requests, request => request.Path == "/opentelemetry.proto.collector.logs.v1.LogsService/Export"
            && request.Payload.Contains("worksheet_read/list", StringComparison.Ordinal)
            && request.Payload.Contains("Outcome", StringComparison.Ordinal));
        Assert.All(requests, request => Assert.Contains("ExcelMcp.Test", request.Payload));
        Assert.All(requests, request => Assert.DoesNotContain("localhost.invalid", request.Payload));
    }

    [Fact]
    public void LocalPipeline_Dispose_DoesNotAddDefaultSdkShutdownWaits()
    {
        var traces = new ShutdownProbe<Activity>();
        var logs = new ShutdownProbe<LogRecord>();
        var telemetry = new LocalTelemetry(
            TelemetrySettings.Resolve(null, key => key == TelemetrySettings.ModeVariable ? "local" : null),
            "ExcelMcp.Test",
            builder => builder.AddProcessor(traces),
            builder => builder.AddProcessor(logs));
        telemetry.Dispose();
        Assert.Equal(0, traces.Timeout);
        Assert.Equal(0, logs.Timeout);
        Assert.True(traces.Disposed);
        Assert.True(logs.Disposed);
    }

    [Theory]
    [InlineData("grpc")]
    [InlineData("http/protobuf")]
    public async Task LocalPipeline_AmbientCredentialHeaders_AreNotSent(string protocol)
    {
        var variables = new Dictionary<string, string>
        {
            ["OTEL_EXPORTER_OTLP_HEADERS"] = "x-test-general=synthetic-general-credential",
            ["OTEL_EXPORTER_OTLP_TRACES_HEADERS"] = "x-test-traces=synthetic-traces-credential",
            ["OTEL_EXPORTER_OTLP_LOGS_HEADERS"] = "x-test-logs=synthetic-logs-credential"
        };
        var previous = variables.Keys.ToDictionary(key => key, Environment.GetEnvironmentVariable);
        try
        {
            foreach (var pair in variables)
                Environment.SetEnvironmentVariable(pair.Key, pair.Value);
            await using var httpReceiver = new LocalOtlpReceiver();
            await using var grpcReceiver = await LocalGrpcOtlpReceiver.StartAsync();
            var settings = TelemetrySettings.Resolve(null, key => key switch
            {
                TelemetrySettings.ModeVariable => "local",
                "OTEL_EXPORTER_OTLP_PROTOCOL" => protocol,
                "OTEL_EXPORTER_OTLP_ENDPOINT" => protocol == "grpc" ? grpcReceiver.Endpoint : httpReceiver.Endpoint,
                _ => Environment.GetEnvironmentVariable(key)
            });
            using var telemetry = new LocalTelemetry(settings, "ExcelMcp.Test");
            telemetry.Track(new EventTelemetry("sheet/list"));
            telemetry.Track(new RequestTelemetry
            {
                Name = "sheet/list",
                Timestamp = DateTimeOffset.UtcNow.AddMilliseconds(-10),
                Duration = TimeSpan.FromMilliseconds(10),
                Success = true,
                ResponseCode = "200"
            });
            Assert.True(telemetry.Flush());
            var requests = protocol == "grpc" ? grpcReceiver.Requests : httpReceiver.Requests;
            Assert.Equal(2, requests.Count);
            Assert.Contains(requests, request => request.Path.Contains(
                protocol == "grpc" ? "TraceService/Export" : "/v1/traces", StringComparison.Ordinal));
            Assert.Contains(requests, request => request.Path.Contains(
                protocol == "grpc" ? "LogsService/Export" : "/v1/logs", StringComparison.Ordinal));
            Assert.All(requests, request =>
            {
                Assert.Contains("sheet/list", request.Payload);
                foreach (var pair in variables)
                {
                    var sentinel = pair.Value.Split('=');
                    Assert.False(request.Headers.ContainsKey(sentinel[0]));
                    Assert.DoesNotContain(request.Headers.Values, value => value.Contains(sentinel[1], StringComparison.Ordinal));
                }
            });
        }
        finally
        {
            foreach (var pair in previous)
                Environment.SetEnvironmentVariable(pair.Key, pair.Value);
        }
    }

    [Fact]
    public async Task LocalPipeline_Redirect_DoesNotForwardTelemetryToAnotherEndpoint()
    {
        await using var destination = new LocalOtlpReceiver();
        await using var receiver = new LocalOtlpReceiver($"{destination.Endpoint}/v1/logs");
        var variables = new Dictionary<string, string>
        {
            [TelemetrySettings.ModeVariable] = "local",
            ["OTEL_EXPORTER_OTLP_ENDPOINT"] = receiver.Endpoint,
            ["OTEL_EXPORTER_OTLP_PROTOCOL"] = "http/protobuf"
        };
        using var telemetry = new LocalTelemetry(
            TelemetrySettings.Resolve(EmbeddedAzure, key => variables.GetValueOrDefault(key)), "ExcelMcp.Test");
        telemetry.Track(new EventTelemetry("sheet/list"));
        Assert.False(telemetry.Flush());
        Assert.Single(receiver.Requests);
        Assert.Empty(destination.Requests);
    }

    private sealed class ShutdownProbe<T> : BaseProcessor<T> where T : class
    {
        internal int? Timeout { get; private set; }
        internal bool Disposed { get; private set; }

        protected override bool OnShutdown(int timeoutMilliseconds)
        {
            Timeout = timeoutMilliseconds;
            return true;
        }

        protected override void Dispose(bool disposing)
        {
            Disposed = true;
            base.Dispose(disposing);
        }
    }
}
