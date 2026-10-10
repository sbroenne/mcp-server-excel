using System.Diagnostics;
using System.Globalization;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Azure.Monitor.OpenTelemetry.Exporter;
using Microsoft.ApplicationInsights.Channel;
using Microsoft.ApplicationInsights.DataContracts;
using OpenTelemetry.Resources;
using Sbroenne.ExcelMcp.CLI.Telemetry;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

// The Azure Monitor exporter keeps process-wide state (shared transmitters, cached role names)
// and these tests change environment variables, so they must not overlap other collections.
[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "Telemetry")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed partial class OpenTelemetryTelemetrySinkTests
{
    private const string StatsbeatVariable = "APPLICATIONINSIGHTS_STATSBEAT_DISABLED";
    private const string SdkStatsVariable = "APPLICATIONINSIGHTS_SDKSTATS_DISABLED";
    private const string SdkDisabledVariable = "OTEL_SDK_DISABLED";

    // The category the Application Insights SDK logged events under. It is stored with every
    // event as the CategoryName property, so it must not change.
    private const string EventCategory = "Microsoft.ApplicationInsights.TelemetryClient";

    private const string MetricExtractorMarker = "(Name: X,Ver:'1.1')";
    private const string TrackPath = "/v2.1/track";

    private static readonly TimeSpan ReceiveTimeout = TimeSpan.FromSeconds(10);

    [GeneratedRegex("^[0-9a-f]{32}$")]
    private static partial Regex OperationIdPattern();

    [GeneratedRegex("^[0-9a-f]{16}$")]
    private static partial Regex SpanIdPattern();

    // Runtime, OpenTelemetry and exporter versions, for example "dotnet10.0.12-servicing.26:otel1.18.0:ext1.8.3".
    [GeneratedRegex(@"^dotnet[^:]+:otel\d+(\.\d+)+:ext(?<exporter>\d+(\.\d+)+)$")]
    private static partial Regex SdkVersionPattern();

    // ---- What reaches Application Insights ----

    [Fact]
    public async Task Track_SucceededCommand_SendsTheEventEnvelope()
    {
        using var endpoint = new FakeIngestionEndpoint();

        var items = await SendAsync(endpoint, "diag.ping", 203, succeeded: true, errorCategory: null);

        var envelope = Assert.Single(endpoint.EnvelopesOfType("EventData"));
        AssertTopLevel(envelope, "Event", "EventData");
        AssertTags(envelope, items.Event, operationName: null);
        var baseData = BaseData(envelope);
        Assert.Equal(2, baseData["ver"]!.GetValue<int>());
        Assert.Equal("diag/ping", baseData["name"]!.GetValue<string>());
        Assert.Equal(
            [
                Pair("Tool", "diag"),
                Pair("Action", "ping"),
                Pair("EntryPoint", "cli"),
                Pair("Success", "True"),
                Pair("Outcome", "succeeded"),
                Pair("DurationMs", "203"),
                Pair("CategoryName", EventCategory)
            ],
            Pairs(baseData["properties"]));
        Assert.Equal(Sorted("ver", "name", "properties"), Keys(baseData));
        var time = DateTime.Parse(envelope["time"]!.GetValue<string>(), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
        Assert.InRange(time, DateTime.UtcNow.AddSeconds(-60), DateTime.UtcNow.AddSeconds(5));
    }

    [Fact]
    public async Task Track_SucceededCommand_SendsTheRequestEnvelope()
    {
        using var endpoint = new FakeIngestionEndpoint();

        var items = await SendAsync(endpoint, "diag.ping", 203, succeeded: true, errorCategory: null);

        var envelope = Assert.Single(endpoint.EnvelopesOfType("RequestData"));
        AssertTopLevel(envelope, "Request", "RequestData");
        AssertTags(envelope, items.Request, operationName: "diag/ping");
        AssertRequestTime(envelope, items.Request);
        var baseData = BaseData(envelope);
        Assert.Equal(2, baseData["ver"]!.GetValue<int>());
        Assert.Matches(SpanIdPattern(), baseData["id"]!.GetValue<string>());
        Assert.Equal("diag/ping", baseData["name"]!.GetValue<string>());
        Assert.Equal("00:00:00.2030000", baseData["duration"]!.GetValue<string>());
        Assert.True(baseData["success"]!.GetValue<bool>());
        Assert.Equal("200", baseData["responseCode"]!.GetValue<string>());
        Assert.Equal(
            [
                Pair("microsoft.sample_rate", "100"),
                Pair("Tool", "diag"),
                Pair("Action", "ping"),
                Pair("EntryPoint", "cli"),
                Pair("Success", "True"),
                Pair("Outcome", "succeeded"),
                Pair("_MS.ProcessedByMetricExtractors", MetricExtractorMarker)
            ],
            Pairs(baseData["properties"]));
        Assert.Equal(Sorted("ver", "id", "name", "duration", "success", "responseCode", "properties"), Keys(baseData));
    }

    // The envelopes below were captured from the real Application Insights SDK for
    // "session close" with an unknown session (error category SessionNotFound).
    [Fact]
    public async Task Track_FailedCommand_SendsTheFailureEventEnvelope()
    {
        using var endpoint = new FakeIngestionEndpoint();

        var items = await SendAsync(endpoint, "session.close", 203, succeeded: false, errorCategory: "SessionNotFound");

        var envelope = Assert.Single(endpoint.EnvelopesOfType("EventData"));
        AssertTags(envelope, items.Event, operationName: null);
        var baseData = BaseData(envelope);
        Assert.Equal("session/close", baseData["name"]!.GetValue<string>());
        Assert.Equal(
            [
                Pair("Tool", "session"),
                Pair("Action", "close"),
                Pair("EntryPoint", "cli"),
                Pair("Success", "False"),
                Pair("Outcome", "failed"),
                Pair("FailureClass", "input-state"),
                Pair("DurationMs", "203"),
                Pair("CategoryName", EventCategory)
            ],
            Pairs(baseData["properties"]));
    }

    [Fact]
    public async Task Track_FailedCommand_SendsTheFailureRequestEnvelope()
    {
        using var endpoint = new FakeIngestionEndpoint();

        var items = await SendAsync(endpoint, "session.close", 203, succeeded: false, errorCategory: "SessionNotFound");

        var envelope = Assert.Single(endpoint.EnvelopesOfType("RequestData"));
        AssertTags(envelope, items.Request, operationName: "session/close");
        AssertRequestTime(envelope, items.Request);
        var baseData = BaseData(envelope);
        Assert.Equal("session/close", baseData["name"]!.GetValue<string>());
        Assert.Equal("00:00:00.2030000", baseData["duration"]!.GetValue<string>());
        Assert.False(baseData["success"]!.GetValue<bool>());
        Assert.Equal("500", baseData["responseCode"]!.GetValue<string>());
        Assert.Equal(
            [
                Pair("microsoft.sample_rate", "100"),
                Pair("Tool", "session"),
                Pair("Action", "close"),
                Pair("EntryPoint", "cli"),
                Pair("Success", "False"),
                Pair("Outcome", "failed"),
                Pair("FailureClass", "input-state"),
                Pair("_MS.ProcessedByMetricExtractors", MetricExtractorMarker)
            ],
            Pairs(baseData["properties"]));
    }

    [Fact]
    public async Task Track_TimedOutCommand_SendsTheFailureCause()
    {
        using var endpoint = new FakeIngestionEndpoint();

        await SendAsync(endpoint, "range.get-values", 1500, succeeded: false, errorCategory: "Timeout");

        var request = Assert.Single(endpoint.EnvelopesOfType("RequestData"));
        var eventEnvelope = Assert.Single(endpoint.EnvelopesOfType("EventData"));
        Assert.Equal("00:00:01.5000000", BaseData(request)["duration"]!.GetValue<string>());
        Assert.Equal(
            [
                Pair("microsoft.sample_rate", "100"),
                Pair("Tool", "range"),
                Pair("Action", "get-values"),
                Pair("EntryPoint", "cli"),
                Pair("Success", "False"),
                Pair("Outcome", "failed"),
                Pair("FailureClass", "timeout-cancellation"),
                Pair("FailureCause", "timeout"),
                Pair("_MS.ProcessedByMetricExtractors", MetricExtractorMarker)
            ],
            Pairs(BaseData(request)["properties"]));
        Assert.Equal(
            [
                Pair("Tool", "range"),
                Pair("Action", "get-values"),
                Pair("EntryPoint", "cli"),
                Pair("Success", "False"),
                Pair("Outcome", "failed"),
                Pair("FailureClass", "timeout-cancellation"),
                Pair("FailureCause", "timeout"),
                Pair("DurationMs", "1500"),
                Pair("CategoryName", EventCategory)
            ],
            Pairs(BaseData(eventEnvelope)["properties"]));
    }

    [Fact]
    public async Task Track_SendsNothingBeyondTheTrackEndpoint()
    {
        using var endpoint = new FakeIngestionEndpoint();

        await SendAsync(endpoint, "diag.ping", 25, succeeded: true, errorCategory: null);

        // No Live Metrics pings, and no other service of the connection string.
        Assert.NotEmpty(endpoint.RequestPaths);
        Assert.All(endpoint.RequestPaths, path => Assert.Equal(TrackPath, path));
    }

    [Fact]
    public async Task Track_SendsStandardMetricsWhenDisposed()
    {
        using var endpoint = new FakeIngestionEndpoint();

        var items = await SendAsync(endpoint, "diag.ping", 203, succeeded: true, errorCategory: null);

        var metrics = endpoint.EnvelopesOfType("MetricData");
        var names = metrics
            .Select(metric => BaseData(metric)["metrics"]![0]!["name"]!.GetValue<string>())
            .ToList();
        // The resource marker, the request duration and five performance counters.
        Assert.Equal(7, names.Count);
        Assert.Contains("_OTELRESOURCE_", names);
        Assert.Contains("requests/duration", names);

        var duration = metrics.Single(metric => BaseData(metric)["metrics"]![0]!["name"]!.GetValue<string>() == "requests/duration");
        var dimensions = Pairs(BaseData(duration)["properties"]).ToDictionary(pair => pair.Key, pair => pair.Value);
        Assert.Equal(items.Request.Context.Cloud.RoleName, dimensions["cloud/roleName"]);
        Assert.Equal(items.Request.Context.Cloud.RoleInstance, dimensions["cloud/roleInstance"]);
        // The standard metric reads the HTTP status code, which a CLI request does not have.
        Assert.Equal("0", dimensions["request/resultCode"]);
        Assert.Equal("True", dimensions["Request.Success"]);
    }

    [Fact]
    public async Task Track_ResourceMarkerCarriesTheIdentityAndDoesNotImpersonateTheApplicationInsightsSdk()
    {
        using var endpoint = new FakeIngestionEndpoint();

        var items = await SendAsync(endpoint, "diag.ping", 25, succeeded: true, errorCategory: null);

        var marker = endpoint.EnvelopesOfType("MetricData")
            .Single(metric => BaseData(metric)["metrics"]![0]!["name"]!.GetValue<string>() == "_OTELRESOURCE_");
        var attributes = Pairs(BaseData(marker)["properties"]).ToDictionary(pair => pair.Key, pair => pair.Value);
        Assert.Equal(items.Request.Context.Cloud.RoleName, attributes["service.name"]);
        Assert.Equal(items.Request.Context.Cloud.RoleInstance, attributes["service.instance.id"]);
        Assert.Equal(items.Request.Context.Component.Version, attributes["service.version"]);
        Assert.DoesNotContain(attributes.Keys, key => key.StartsWith("telemetry.distro.", StringComparison.Ordinal));
    }

    [Fact]
    public async Task FlushAsync_BeforeDispose_DeliversTheEventAndTheRequest()
    {
        using var endpoint = new FakeIngestionEndpoint();
        var items = CliTelemetry.CreateCommandInvocationTelemetry("diag.ping", 25, succeeded: true, errorCategory: null);
        var sink = CreateSink(endpoint);
        try
        {
            sink.Track(items.Event, items.Request);

            await sink.FlushAsync();

            Assert.Single(endpoint.EnvelopesOfType("EventData"));
            Assert.Single(endpoint.EnvelopesOfType("RequestData"));
        }
        finally
        {
            sink.Dispose();
        }
    }

    // ---- Things that must hold whatever the caller is doing ----

    [Fact]
    public async Task Track_InsideAnAmbientActivity_StillSendsARootRequestAndAnEventWithoutOperationIds()
    {
        using var endpoint = new FakeIngestionEndpoint();
        var items = CliTelemetry.CreateCommandInvocationTelemetry("diag.ping", 25, succeeded: true, errorCategory: null);
        var sink = CreateSink(endpoint);
        var ambient = new Activity("ambient-operation").Start();
        try
        {
            sink.Track(items.Event, items.Request);

            // The caller's Activity is left exactly as it was.
            Assert.Same(ambient, Activity.Current);
            await sink.FlushAsync();
        }
        finally
        {
            ambient.Stop();
            sink.Dispose();
        }

        var request = Assert.Single(endpoint.EnvelopesOfType("RequestData"));
        var eventEnvelope = Assert.Single(endpoint.EnvelopesOfType("EventData"));
        var requestTags = request["tags"]!.AsObject();
        Assert.False(requestTags.ContainsKey("ai.operation.parentId"));
        Assert.NotEqual(ambient.TraceId.ToHexString(), requestTags["ai.operation.id"]!.GetValue<string>());
        var eventTags = eventEnvelope["tags"]!.AsObject();
        Assert.False(eventTags.ContainsKey("ai.operation.id"));
        Assert.False(eventTags.ContainsKey("ai.operation.parentId"));
    }

    [Fact]
    public async Task Track_SpanStartedRightAfterTheSdkIsBuilt_IsAlwaysKept()
    {
        // The exporter's default sampler keeps only a fraction of the spans in the first
        // 200 ms of a process, which is exactly when a quick CLI command reports.
        for (var attempt = 0; attempt < 5; attempt++)
        {
            using var endpoint = new FakeIngestionEndpoint();

            await SendAsync(endpoint, "diag.ping", 5, succeeded: true, errorCategory: null);

            var request = Assert.Single(endpoint.EnvelopesOfType("RequestData"));
            Assert.Equal("100", Pairs(BaseData(request)["properties"]).First(pair => pair.Key == "microsoft.sample_rate").Value);
        }
    }

    [Fact]
    public async Task Track_ManyCallsInARow_SendsEveryRecord()
    {
        using var endpoint = new FakeIngestionEndpoint();
        var sink = CreateSink(endpoint);
        try
        {
            for (var i = 0; i < 25; i++)
            {
                var items = CliTelemetry.CreateCommandInvocationTelemetry("range.get-values", i, succeeded: true, errorCategory: null);
                sink.Track(items.Event, items.Request);
            }

            await sink.FlushAsync();
        }
        finally
        {
            sink.Dispose();
        }

        Assert.Equal(25, endpoint.EnvelopesOfType("EventData").Count);
        var requests = endpoint.EnvelopesOfType("RequestData");
        Assert.Equal(25, requests.Count);
        Assert.All(requests, request =>
            Assert.Equal("100", Pairs(BaseData(request)["properties"]).First(pair => pair.Key == "microsoft.sample_rate").Value));
    }

    [Fact]
    public async Task Track_DoesNotChangeTheTelemetryItems()
    {
        using var endpoint = new FakeIngestionEndpoint();
        var items = CliTelemetry.CreateCommandInvocationTelemetry("diag.ping", 25, succeeded: true, errorCategory: null);
        var eventProperties = items.Event.Properties.ToList();
        var requestProperties = items.Request.Properties.ToList();
        var sink = CreateSink(endpoint);
        try
        {
            sink.Track(items.Event, items.Request);
            await sink.FlushAsync();
        }
        finally
        {
            sink.Dispose();
        }

        Assert.Equal(eventProperties, items.Event.Properties.ToList());
        Assert.Equal(requestProperties, items.Request.Properties.ToList());
    }

    [Fact]
    public void Dispose_CalledTwice_DoesNotThrow()
    {
        using var endpoint = new FakeIngestionEndpoint();
        var sink = CreateSink(endpoint);

        sink.Dispose();
        var exception = Record.Exception(sink.Dispose);

        Assert.Null(exception);
    }

    [Fact]
    public async Task UseAfterDispose_DoesNotThrow()
    {
        using var endpoint = new FakeIngestionEndpoint();
        var items = CliTelemetry.CreateCommandInvocationTelemetry("diag.ping", 25, succeeded: true, errorCategory: null);
        var sink = CreateSink(endpoint);
        sink.Dispose();

        var trackException = Record.Exception(() => sink.Track(items.Event, items.Request));
        var flushException = await Record.ExceptionAsync(sink.FlushAsync);

        Assert.Null(trackException);
        Assert.Null(flushException);
    }

    [Fact]
    public async Task Track_NullItems_DoesNotThrow()
    {
        using var endpoint = new FakeIngestionEndpoint();
        var sink = CreateSink(endpoint);
        try
        {
            var trackException = Record.Exception(() => sink.Track(null!, null!));
            var flushException = await Record.ExceptionAsync(sink.FlushAsync);

            Assert.Null(trackException);
            Assert.Null(flushException);
        }
        finally
        {
            sink.Dispose();
        }
    }

    [Fact]
    public async Task Track_EndpointUnreachable_DoesNotThrow()
    {
        // Nothing listens on the port once the fake endpoint is gone.
        string connectionString;
        using (var endpoint = new FakeIngestionEndpoint())
        {
            connectionString = endpoint.ConnectionString;
        }

        var items = CliTelemetry.CreateCommandInvocationTelemetry("diag.ping", 25, succeeded: true, errorCategory: null);
        var sink = OpenTelemetryTelemetrySink.Create(connectionString, CliTelemetry.Identity, disableOfflineStorage: true);

        var trackException = Record.Exception(() => sink.Track(items.Event, items.Request));
        var flushException = await Record.ExceptionAsync(sink.FlushAsync);
        var disposeException = Record.Exception(sink.Dispose);

        Assert.Null(trackException);
        Assert.Null(flushException);
        Assert.Null(disposeException);
    }

    // ---- OTEL_SDK_DISABLED ----

    [Theory]
    [InlineData("true")]
    [InlineData("TRUE")]
    [InlineData("True")]
    public void IsSdkDisabled_TrueInAnyCase_IsDisabled(string value) =>
        Assert.True(OpenTelemetryTelemetrySink.IsSdkDisabled(_ => value));

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("false")]
    [InlineData("0")]
    [InlineData("yes")]
    public void IsSdkDisabled_AnythingElse_IsEnabled(string? value) =>
        Assert.False(OpenTelemetryTelemetrySink.IsSdkDisabled(_ => value));

    [Fact]
    public async Task OtelSdkDisabled_SendsNothingAndDoesNotThrow()
    {
        using var endpoint = new FakeIngestionEndpoint();
        var items = CliTelemetry.CreateCommandInvocationTelemetry("diag.ping", 25, succeeded: true, errorCategory: null);
        var previous = Environment.GetEnvironmentVariable(SdkDisabledVariable);
        Environment.SetEnvironmentVariable(SdkDisabledVariable, "true");
        try
        {
            var sink = CreateSink(endpoint);
            var trackException = Record.Exception(() => sink.Track(items.Event, items.Request));
            var flushException = await Record.ExceptionAsync(sink.FlushAsync);
            var disposeException = Record.Exception(sink.Dispose);

            Assert.Null(trackException);
            Assert.Null(flushException);
            Assert.Null(disposeException);
        }
        finally
        {
            Environment.SetEnvironmentVariable(SdkDisabledVariable, previous);
        }

        Assert.Empty(endpoint.RequestPaths);
        Assert.Empty(endpoint.Envelopes);
    }

    // ---- Set-up ----

    [Fact]
    public void Create_TurnsOffTheSdksOwnHealthReports()
    {
        using var endpoint = new FakeIngestionEndpoint();
        var previousStatsbeat = Environment.GetEnvironmentVariable(StatsbeatVariable);
        var previousSdkStats = Environment.GetEnvironmentVariable(SdkStatsVariable);
        Environment.SetEnvironmentVariable(StatsbeatVariable, null);
        Environment.SetEnvironmentVariable(SdkStatsVariable, null);
        try
        {
            CreateSink(endpoint).Dispose();

            Assert.Equal("true", Environment.GetEnvironmentVariable(StatsbeatVariable));
            Assert.Equal("true", Environment.GetEnvironmentVariable(SdkStatsVariable));
        }
        finally
        {
            Environment.SetEnvironmentVariable(StatsbeatVariable, previousStatsbeat);
            Environment.SetEnvironmentVariable(SdkStatsVariable, previousSdkStats);
        }
    }

    [Fact]
    public void ConfigureResource_SetsTheRoleNameInstanceAndVersion()
    {
        var identity = new TelemetryIdentity("Some.Role", "instance-0123abcd", "9.8.7+abc");

        var resource = OpenTelemetryTelemetrySink.ConfigureResource(ResourceBuilder.CreateEmpty(), identity).Build();

        var attributes = resource.Attributes.ToDictionary(pair => pair.Key, pair => pair.Value);
        Assert.Equal("Some.Role", attributes["service.name"]);
        Assert.Equal("instance-0123abcd", attributes["service.instance.id"]);
        Assert.Equal("9.8.7+abc", attributes["service.version"]);
    }

    [Fact]
    public void ConfigureResource_DoesNotAddDistroAttributes()
    {
        var identity = new TelemetryIdentity("Some.Role", "instance-0123abcd", "9.8.7+abc");

        var resource = OpenTelemetryTelemetrySink.ConfigureResource(ResourceBuilder.CreateEmpty(), identity).Build();

        Assert.DoesNotContain(resource.Attributes, pair => pair.Key.StartsWith("telemetry.distro.", StringComparison.Ordinal));
    }

    [Fact]
    public void CliTelemetryIdentity_IsWhatTheTelemetryItemsCarry()
    {
        var items = CliTelemetry.CreateCommandInvocationTelemetry("diag.ping", 25, succeeded: true, errorCategory: null);

        foreach (var context in new[] { items.Event.Context, items.Request.Context })
        {
            Assert.Equal(CliTelemetry.Identity.RoleName, context.Cloud.RoleName);
            Assert.Equal(CliTelemetry.Identity.RoleInstance, context.Cloud.RoleInstance);
            Assert.Equal(CliTelemetry.Identity.Version, context.Component.Version);
        }

        Assert.Equal($"instance-{items.Request.Context.User.Id[..8]}", CliTelemetry.Identity.RoleInstance);
    }

    [Fact]
    public void DisableSdkSelfTelemetry_VariablesUnset_SetsBothSwitches()
    {
        var variables = new Dictionary<string, string>();

        OpenTelemetryTelemetrySink.DisableSdkSelfTelemetry(
            name => variables.GetValueOrDefault(name),
            (name, value) => variables[name] = value);

        Assert.Equal("true", variables[StatsbeatVariable]);
        Assert.Equal("true", variables[SdkStatsVariable]);
    }

    [Fact]
    public void DisableSdkSelfTelemetry_VariablesEmpty_SetsBothSwitches()
    {
        var variables = new Dictionary<string, string>
        {
            [StatsbeatVariable] = "",
            [SdkStatsVariable] = ""
        };

        OpenTelemetryTelemetrySink.DisableSdkSelfTelemetry(
            name => variables.GetValueOrDefault(name),
            (name, value) => variables[name] = value);

        Assert.Equal("true", variables[StatsbeatVariable]);
        Assert.Equal("true", variables[SdkStatsVariable]);
    }

    [Theory]
    [InlineData(StatsbeatVariable)]
    [InlineData(SdkStatsVariable)]
    public void DisableSdkSelfTelemetry_ExplicitValue_IsLeftUntouched(string explicitVariable)
    {
        var otherVariable = explicitVariable == StatsbeatVariable ? SdkStatsVariable : StatsbeatVariable;
        var variables = new Dictionary<string, string>
        {
            [explicitVariable] = "false"
        };
        var written = new List<string>();

        OpenTelemetryTelemetrySink.DisableSdkSelfTelemetry(
            name => variables.GetValueOrDefault(name),
            (name, value) =>
            {
                written.Add(name);
                variables[name] = value;
            });

        Assert.Equal("false", variables[explicitVariable]);
        Assert.Equal("true", variables[otherVariable]);
        Assert.Equal([otherVariable], written);
    }

    // ---- Helpers ----

    // Offline storage would write unsent telemetry under the user's profile.
    private static OpenTelemetryTelemetrySink CreateSink(FakeIngestionEndpoint endpoint) =>
        OpenTelemetryTelemetrySink.Create(endpoint.ConnectionString, CliTelemetry.Identity, disableOfflineStorage: true);

    /// <summary>Builds the items for one command, sends them through a new sink, and waits for the sink to shut down.</summary>
    private static async Task<(EventTelemetry Event, RequestTelemetry Request)> SendAsync(
        FakeIngestionEndpoint endpoint,
        string command,
        long durationMs,
        bool succeeded,
        string? errorCategory)
    {
        var items = CliTelemetry.CreateCommandInvocationTelemetry(command, durationMs, succeeded, errorCategory);
        var sink = CreateSink(endpoint);
        try
        {
            sink.Track(items.Event, items.Request);
            await sink.FlushAsync();
        }
        finally
        {
            sink.Dispose();
        }

        // Dispose exports the last batches, including the standard metrics.
        Assert.True(
            endpoint.WaitFor(
                envelopes => envelopes.Any(e => FakeIngestionEndpoint.BaseType(e) == "RequestData")
                    && envelopes.Any(e => FakeIngestionEndpoint.BaseType(e) == "EventData")
                    && envelopes.Count(e => FakeIngestionEndpoint.BaseType(e) == "MetricData") >= 7,
                ReceiveTimeout),
            "The expected telemetry did not reach the endpoint.");
        return items;
    }

    private static string[] Sorted(params string[] names) => [.. names.Order(StringComparer.Ordinal)];

    private static string[] Keys(JsonObject node) => [.. node.Select(entry => entry.Key).Order(StringComparer.Ordinal)];

    private static KeyValuePair<string, string> Pair(string key, string value) => new(key, value);

    private static List<KeyValuePair<string, string>> Pairs(JsonNode? node) =>
        node is JsonObject obj
            ? [.. obj.Select(entry => Pair(entry.Key, entry.Value!.GetValue<string>()))]
            : [];

    private static JsonObject BaseData(JsonObject envelope) => envelope["data"]!["baseData"]!.AsObject();

    private static void AssertTopLevel(JsonObject envelope, string name, string baseType)
    {
        Assert.Equal(name, envelope["name"]!.GetValue<string>());
        Assert.Equal(FakeIngestionEndpoint.InstrumentationKey, envelope["iKey"]!.GetValue<string>());
        Assert.Equal(baseType, envelope["data"]!["baseType"]!.GetValue<string>());
        // No envelope version, sample rate or sequence number.
        Assert.Equal(Sorted("name", "time", "iKey", "tags", "data"), Keys(envelope));
    }

    /// <summary>Checks the tags of an envelope; a request (operation name given) also carries an operation id and name.</summary>
    private static void AssertTags(JsonObject envelope, ITelemetry source, string? operationName)
    {
        var tags = Pairs(envelope["tags"]).ToDictionary(pair => pair.Key, pair => pair.Value);

        var sdkVersion = tags["ai.internal.sdkVersion"];
        var match = SdkVersionPattern().Match(sdkVersion);
        Assert.True(match.Success, $"Unexpected ai.internal.sdkVersion: {sdkVersion}");
        var exporterVersion = typeof(AzureMonitorExporterOptions).Assembly.GetName().Version!;
        Assert.Equal($"{exporterVersion.Major}.{exporterVersion.Minor}.{exporterVersion.Build}", match.Groups["exporter"].Value);
        tags.Remove("ai.internal.sdkVersion");

        var expected = new Dictionary<string, string>
        {
            ["ai.user.id"] = source.Context.User.Id,
            ["ai.session.id"] = source.Context.Session.Id,
            ["ai.cloud.role"] = source.Context.Cloud.RoleName,
            ["ai.cloud.roleInstance"] = source.Context.Cloud.RoleInstance,
            ["ai.application.ver"] = source.Context.Component.Version
        };
        if (operationName != null)
        {
            // The operation id is random.
            Assert.Matches(OperationIdPattern(), tags["ai.operation.id"]);
            tags.Remove("ai.operation.id");
            expected["ai.operation.name"] = operationName;
        }

        Assert.Equal(expected, tags);
        Assert.Equal("ExcelMcp.CLI", tags["ai.cloud.role"]);
        Assert.Matches("^instance-[0-9a-f]{8}$", tags["ai.cloud.roleInstance"]);
        Assert.Matches("^[0-9a-f]{16}$", tags["ai.user.id"]);
        Assert.Matches("^[0-9a-f]{8}$", tags["ai.session.id"]);
    }

    private static void AssertRequestTime(JsonObject envelope, RequestTelemetry request)
    {
        var time = DateTime.Parse(envelope["time"]!.GetValue<string>(), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
        Assert.Equal(request.Timestamp.UtcDateTime, time);
    }
}
