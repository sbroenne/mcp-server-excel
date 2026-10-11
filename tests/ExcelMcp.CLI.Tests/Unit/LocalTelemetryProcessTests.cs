using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.Tests.Telemetry;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "Telemetry")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class LocalTelemetryProcessTests
{
    [Theory]
    [InlineData("local")]
    [InlineData("off")]
    public async Task Main_SelectedMode_PreservesJsonAndFlushesOnlyLocalTelemetry(string mode)
    {
        await using var receiver = new LocalOtlpReceiver();
        var start = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        start.ArgumentList.Add(typeof(Program).Assembly.Location);
        start.ArgumentList.Add("service");
        start.ArgumentList.Add("status");
        start.ArgumentList.Add("--quiet");
        start.Environment["EXCELMCP_TELEMETRY_MODE"] = mode;
        start.Environment["OTEL_EXPORTER_OTLP_ENDPOINT"] = receiver.Endpoint;
        start.Environment["OTEL_EXPORTER_OTLP_PROTOCOL"] = "http/protobuf";
        start.Environment["OTEL_RESOURCE_ATTRIBUTES"] = "private_path=confidential-workbook.xlsx";
        start.Environment["EXCELMCP_CLI_PIPE"] = $"local-telemetry-test-{Guid.NewGuid():N}";
        using var process = new Process { StartInfo = start };
        Assert.True(process.Start());
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        try
        {
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(30));
        }
        finally
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
                await process.WaitForExitAsync();
            }
        }
        Assert.Equal(0, process.ExitCode);
        using var result = JsonDocument.Parse(await stdout);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.DoesNotContain("Local telemetry", await stderr);
        var requests = receiver.Requests;
        if (mode == "off")
        {
            Assert.Empty(requests);
        }
        else
        {
            Assert.Equal(2, requests.Count);
            Assert.Contains(requests, request => request.Path == "/v1/traces" && request.Payload.Contains("service/status", StringComparison.Ordinal));
            Assert.Contains(requests, request => request.Path == "/v1/logs" && request.Payload.Contains("service/status", StringComparison.Ordinal));
            Assert.All(requests, request => Assert.DoesNotContain("confidential-workbook.xlsx", request.Payload));
        }
    }
}
