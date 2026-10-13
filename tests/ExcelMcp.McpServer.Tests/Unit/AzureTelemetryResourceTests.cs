using System.Reflection;
using System.Text.Json;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.WorkerService;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Sbroenne.ExcelMcp.Tests.Telemetry;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Collection("ProgramTransport")]
[Trait("Layer", "McpServer")]
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
            var builder = Host.CreateApplicationBuilder();
            Program.ConfigureTelemetry(builder, TelemetrySettings.Resolve(
                $"InstrumentationKey=00000000-0000-0000-000000000000;IngestionEndpoint={receiver.Endpoint}/",
                _ => null));
            builder.Services.Configure<ApplicationInsightsServiceOptions>(options =>
            {
                options.TracesPerSecond = 1_000_000;
            });
            Program.ConfigureStdioLogging(builder.Logging);
            using var host = builder.Build();
            await host.StartAsync();
            var client = host.Services.GetRequiredService<TelemetryClient>();
            ExcelMcpTelemetry.InitializeUserIdentity();
            ExcelMcpTelemetry.SetTelemetryClient(client);
            var (eventTelemetry, request) = ExcelMcpTelemetry.CreateToolInvocationTelemetry(
                "file_read", "list", 123, new ToolInvocationResult(ToolInvocationOutcome.Succeeded, null));
            client.TrackEvent(eventTelemetry);
            client.TrackRequest(request);
            Assert.True(await client.FlushAsync(CancellationToken.None).WaitAsync(TimeSpan.FromSeconds(15)));
            await host.StopAsync();
            host.Dispose();

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
                Assert.True(attributes.GetProperty("service.instance.id").GetString() == request.Context.Cloud.RoleInstance,
                    "The exported resource must use the same anonymous instance as tool records.");
                Assert.Equal("ExcelMcp.McpServer", attributes.GetProperty("service.name").GetString());
                Assert.Equal(version, attributes.GetProperty("service.version").GetString());
            });
            var metric = Assert.Single(items, item =>
                item.GetProperty("data").GetProperty("baseType").GetString() == "MetricData" &&
                item.GetProperty("data").GetProperty("baseData").GetProperty("metrics")
                    .EnumerateArray().Any(value => value.GetProperty("name").GetString() == "requests/duration"));
            var properties = metric.GetProperty("data").GetProperty("baseData").GetProperty("properties");
            Assert.True(properties.GetProperty("cloud/roleInstance").GetString() == request.Context.Cloud.RoleInstance,
                "Duration metrics must use the same anonymous instance as tool records.");
            Assert.Equal("ExcelMcp.McpServer", properties.GetProperty("cloud/roleName").GetString());
            var record = Assert.Single(items, item =>
                item.GetProperty("data").GetProperty("baseType").GetString() == "RequestData" &&
                item.GetProperty("data").GetProperty("baseData").GetProperty("name").GetString() == "file_read/list");
            var tags = record.GetProperty("tags");
            Assert.True(tags.GetProperty("ai.cloud.roleInstance").GetString() == request.Context.Cloud.RoleInstance,
                "Tool records must retain their anonymous instance.");
            Assert.Equal("ExcelMcp.McpServer", tags.GetProperty("ai.cloud.role").GetString());
            Assert.Equal(version, tags.GetProperty("ai.application.ver").GetString());
        }
        finally
        {
            ExcelMcpTelemetry.SetLocalTelemetry(null);
            foreach (var pair in previous)
                Environment.SetEnvironmentVariable(pair.Key, pair.Value);
        }
    }
}
