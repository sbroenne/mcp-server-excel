using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
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
    [Fact]
    public async Task Daemon_ExportsCommandAfterShortLivedClientExits()
    {
        await using var receiver = new LocalOtlpReceiver();
        var pipe = $"daemon-telemetry-test-{Guid.NewGuid():N}";
        var start = CreateStartInfo(pipe, receiver.Endpoint, "service", "run", "--pipe-name", pipe);
        using var daemon = Process.Start(start);
        Assert.NotNull(daemon);
        var daemonOutput = daemon.StandardOutput.ReadToEndAsync();
        var daemonError = daemon.StandardError.ReadToEndAsync();
        using var client = new ServiceClient(pipe, TimeSpan.FromMilliseconds(200), TimeSpan.FromSeconds(5));
        try
        {
            var deadline = DateTime.UtcNow.AddSeconds(30);
            while (true)
            {
                Assert.False(daemon.HasExited);
                var ping = await client.SendAsync(new ServiceRequest { Command = "service.ping" });
                if (ping.Success) break;
                Assert.True(DateTime.UtcNow < deadline, "The private daemon did not become ready.");
                await Task.Delay(100);
            }

            using var command = Process.Start(CreateStartInfo(pipe, receiver.Endpoint, "session", "list", "--quiet"));
            Assert.NotNull(command);
            var output = command.StandardOutput.ReadToEndAsync();
            var error = command.StandardError.ReadToEndAsync();
            try
            {
                await command.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(15));
                Assert.Equal(0, command.ExitCode);
                using var result = JsonDocument.Parse(await output);
                Assert.True(result.RootElement.GetProperty("success").GetBoolean());
                Assert.Empty(await error);
            }
            finally
            {
                if (!command.HasExited)
                {
                    command.Kill(entireProcessTree: true);
                    await command.WaitForExitAsync();
                }
            }

            deadline = DateTime.UtcNow.AddSeconds(15);
            while (!receiver.Requests.Any(r => r.Path == "/v1/logs" && r.Payload.Contains("session/list")) ||
                   !receiver.Requests.Any(r => r.Path == "/v1/traces" && r.Payload.Contains("session/list")))
            {
                Assert.True(DateTime.UtcNow < deadline,
                    "The daemon must export both the event and operation after the client exits.");
                await Task.Delay(100);
            }
            Assert.False(daemon.HasExited);
            Assert.All(receiver.Requests, r => Assert.DoesNotContain("service/ping", r.Payload));
        }
        finally
        {
            if (!daemon.HasExited)
            {
                await client.SendAsync(new ServiceRequest { Command = "service.shutdown" });
                try
                {
                    await daemon.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
                }
                finally
                {
                    if (!daemon.HasExited)
                    {
                        daemon.Kill(entireProcessTree: true);
                        await daemon.WaitForExitAsync();
                    }
                }
            }
            await daemonOutput;
            await daemonError;
        }
    }

    private static ProcessStartInfo CreateStartInfo(string pipe, string endpoint, params string[] args)
    {
        var start = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        start.ArgumentList.Add(typeof(Program).Assembly.Location);
        foreach (var arg in args) start.ArgumentList.Add(arg);
        start.Environment["EXCELMCP_TELEMETRY_MODE"] = "local";
        start.Environment["OTEL_EXPORTER_OTLP_ENDPOINT"] = endpoint;
        start.Environment["OTEL_EXPORTER_OTLP_PROTOCOL"] = "http/protobuf";
        start.Environment["EXCELMCP_CLI_PIPE"] = pipe;
        return start;
    }

    [Theory]
    [InlineData("local")]
    [InlineData("off")]
    public async Task Main_SelectedMode_PreservesJsonWithoutWaitingForReporting(string mode)
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
        // Local reporting is best effort: a short command may exit before any export.
        Assert.All(requests, request => Assert.DoesNotContain("confidential-workbook.xlsx", request.Payload));
    }
}
