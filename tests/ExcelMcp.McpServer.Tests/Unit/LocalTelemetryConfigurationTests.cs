using Microsoft.Extensions.Hosting;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Tests.Telemetry;
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
    public void ConfigureTelemetry_LocalMode_RegistersLocalDeliveryWithoutAzure()
    {
        var variables = new Dictionary<string, string?>
        {
            [TelemetrySettings.ModeVariable] = "local",
            ["OTEL_EXPORTER_OTLP_ENDPOINT"] = null,
            ["OTEL_EXPORTER_OTLP_PROTOCOL"] = null
        };
        var previous = variables.Keys.ToDictionary(key => key, Environment.GetEnvironmentVariable);
        try
        {
            foreach (var pair in variables)
                Environment.SetEnvironmentVariable(pair.Key, pair.Value);
            var builder = Host.CreateApplicationBuilder();
            Program.ConfigureTelemetry(builder);

            Assert.Contains(builder.Services, service => service.ServiceType.Name == "LocalTelemetry");
            Assert.DoesNotContain(builder.Services, service =>
                service.ServiceType.FullName == "Microsoft.ApplicationInsights.TelemetryClient");
        }
        finally
        {
            foreach (var pair in previous)
                Environment.SetEnvironmentVariable(pair.Key, pair.Value);
        }
    }

    [Theory]
    [InlineData("local")]
    [InlineData("off")]
    public void ConfigureTelemetry_OverrideEmbeddedAzure_DoesNotRegisterAzure(string mode)
    {
        var settings = TelemetrySettings.Resolve(
            "InstrumentationKey=00000000-0000-0000-0000-000000000000;IngestionEndpoint=https://localhost.invalid/",
            key => key == TelemetrySettings.ModeVariable ? mode : null);
        var builder = Host.CreateApplicationBuilder();
        Program.ConfigureTelemetry(builder, settings);
        Assert.DoesNotContain(builder.Services, service =>
            service.ServiceType.FullName?.StartsWith("Microsoft.ApplicationInsights.", StringComparison.Ordinal) == true);
        Assert.Equal(mode == "local", builder.Services.Any(service => service.ServiceType == typeof(LocalTelemetry)));
    }

    [Fact]
    public async Task RunAsync_LocalMode_DeliversProtocolInvocationAndSanitizedExceptionOnShutdown()
    {
        await using var receiver = new LocalOtlpReceiver();
        var variables = new Dictionary<string, string>
        {
            [TelemetrySettings.ModeVariable] = "local",
            ["OTEL_EXPORTER_OTLP_ENDPOINT"] = receiver.Endpoint,
            ["OTEL_EXPORTER_OTLP_PROTOCOL"] = "http/protobuf",
            ["OTEL_RESOURCE_ATTRIBUTES"] = "private_path=confidential-workbook.xlsx"
        };
        var previous = variables.Keys.ToDictionary(key => key, Environment.GetEnvironmentVariable);
        var fixture = new RecordingProgramTransportFixture();
        try
        {
            foreach (var pair in variables)
                Environment.SetEnvironmentVariable(pair.Key, pair.Value);
            await fixture.InitializeAsync();
            await fixture.CallToolAsync(
                "file_read", new Dictionary<string, object?> { ["action"] = "list" },
                new ServiceResponse { Success = true, Result = """{"success":true,"sessions":[]}""" },
                "session.list", expectedSessionId: null, expectedArgsJson: null);
            ExcelMcpTelemetry.TrackUnhandledException(
                new InvalidOperationException("confidential-workbook.xlsx customer@example.com"),
                "McpServer.RunAsync");
        }
        finally
        {
            try
            {
                await fixture.DisposeAsync();
            }
            finally
            {
                foreach (var pair in previous)
                    Environment.SetEnvironmentVariable(pair.Key, pair.Value);
            }
        }

        var requests = receiver.Requests;
        Assert.Contains(requests, request => request.Path == "/v1/traces" && request.Payload.Contains("file_read/list", StringComparison.Ordinal));
        Assert.Contains(requests, request => request.Path == "/v1/logs" && request.Payload.Contains("SessionStart", StringComparison.Ordinal));
        Assert.Contains(requests, request => request.Path == "/v1/logs" && request.Payload.Contains("UnhandledException", StringComparison.Ordinal)
            && request.Payload.Contains("Sanitized", StringComparison.Ordinal));
        Assert.All(requests, request =>
        {
            Assert.DoesNotContain("confidential-workbook.xlsx", request.Payload);
            Assert.DoesNotContain("customer@example.com", request.Payload);
            Assert.DoesNotContain("RecordingProgramTransportClient", request.Payload);
        });
    }
}
