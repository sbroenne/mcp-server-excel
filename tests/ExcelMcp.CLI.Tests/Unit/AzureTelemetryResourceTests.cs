using System.Reflection;
using System.Text.Json;
using Microsoft.ApplicationInsights;
using Sbroenne.ExcelMcp.CLI.Telemetry;
using Sbroenne.ExcelMcp.Telemetry;
using Sbroenne.ExcelMcp.Tests.Telemetry;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "Telemetry")]
[Trait("Speed", "Slow")]
[Trait("RequiresExcel", "false")]
public sealed class AzureTelemetryResourceTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("synthetic-computer-name")]
    public async Task AzureConfiguration_ExportsAnonymousMetricAndRecordIdentity(string? ambientInstance)
    {
        await using var receiver = new LocalOtlpReceiver(responseBody:
            """{"itemsReceived":2,"itemsAccepted":2,"errors":[]}""");
        var variables = new Dictionary<string, string?>
        {
            ["APPLICATIONINSIGHTS_STATSBEAT_DISABLED"] = "true",
            ["APPLICATIONINSIGHTS_SDKSTATS_DISABLED"] = "true",
            ["OTEL_RESOURCE_ATTRIBUTES"] = ambientInstance is null ? null
                : $"service.instance.id={ambientInstance}",
            ["OTEL_SERVICE_NAME"] = "synthetic-service"
        };
        var previous = variables.Keys.ToDictionary(key => key, Environment.GetEnvironmentVariable);
        try
        {
            foreach (var pair in variables)
                Environment.SetEnvironmentVariable(pair.Key, pair.Value);
            using var identity = CliTelemetry.UseUserIdentityFactoryForTesting(() =>
                new TelemetryUserIdentity("0123456789abcdef0123456789abcdef", null));
            using var configuration = ApplicationInsightsTelemetrySink.CreateConfiguration(
                $"InstrumentationKey=00000000-0000-0000-0000-000000000000;IngestionEndpoint={receiver.Endpoint}/");
            configuration.TracesPerSecond = 1_000_000;
            configuration.DisableOfflineStorage = true;
            var client = new TelemetryClient(configuration);
            var (eventTelemetry, request) = CliTelemetry.CreateCommandInvocationTelemetry(
                "diag.ping", 123, succeeded: true, errorCategory: null);
            client.TrackEvent(eventTelemetry);
            client.TrackRequest(request);
            Assert.True(await client.FlushAsync(CancellationToken.None).WaitAsync(TimeSpan.FromSeconds(15)));
            configuration.Dispose();

            var items = receiver.Requests.SelectMany(capture =>
            {
                Assert.Equal("/v2.1/track", capture.Path);
                Assert.False(capture.Payload.Contains(Environment.MachineName, StringComparison.OrdinalIgnoreCase),
                    "Exported telemetry must not contain the computer name.");
                if (ambientInstance is not null)
                    Assert.DoesNotContain(ambientInstance, capture.Payload);
                return capture.Payload.Split('\n', StringSplitOptions.RemoveEmptyEntries)
                    .Select(line => JsonSerializer.Deserialize<JsonElement>(line));
            }).ToArray();
            var version = typeof(Program).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()!
                .InformationalVersion;
            var resources = items.Where(item =>
                item.GetProperty("data").GetProperty("baseType").GetString() == "MetricData" &&
                item.GetProperty("data").GetProperty("baseData").GetProperty("metrics")
                    .EnumerateArray().Any(value => value.GetProperty("name").GetString() == "_OTELRESOURCE_")).ToArray();
            Assert.NotEmpty(resources);
            Assert.All(resources, resource =>
            {
                var attributes = resource.GetProperty("data").GetProperty("baseData").GetProperty("properties");
                Assert.Equal("instance-01234567", attributes.GetProperty("service.instance.id").GetString());
                Assert.Equal("ExcelMcp.CLI", attributes.GetProperty("service.name").GetString());
                Assert.Equal(version, attributes.GetProperty("service.version").GetString());
            });
            var metric = Assert.Single(items, item =>
                item.GetProperty("data").GetProperty("baseType").GetString() == "MetricData" &&
                item.GetProperty("data").GetProperty("baseData").GetProperty("metrics")
                    .EnumerateArray().Any(value => value.GetProperty("name").GetString() == "requests/duration"));
            var properties = metric.GetProperty("data").GetProperty("baseData").GetProperty("properties");
            Assert.True(properties.GetProperty("cloud/roleInstance").GetString() == "instance-01234567",
                "Duration metrics must use the same anonymous instance as command records.");
            Assert.Equal("ExcelMcp.CLI", properties.GetProperty("cloud/roleName").GetString());
            foreach (var baseType in new[] { "EventData", "RequestData" })
            {
                var record = Assert.Single(items, item =>
                    item.GetProperty("data").GetProperty("baseType").GetString() == baseType &&
                    item.GetProperty("data").GetProperty("baseData").GetProperty("name").GetString() == "diag/ping");
                var tags = record.GetProperty("tags");
                Assert.Equal("instance-01234567", tags.GetProperty("ai.cloud.roleInstance").GetString());
                Assert.Equal("ExcelMcp.CLI", tags.GetProperty("ai.cloud.role").GetString());
                Assert.Equal(version, tags.GetProperty("ai.application.ver").GetString());
            }
        }
        finally
        {
            foreach (var pair in previous)
                Environment.SetEnvironmentVariable(pair.Key, pair.Value);
        }
    }
}
