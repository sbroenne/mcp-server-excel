using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[CollectionDefinition("Mac VBA helper environment", DisableParallelization = true)]
public sealed class MacVbaHelperEnvironmentFixture;

[Collection("Mac VBA helper environment")]
public sealed class MacVbaHelperClientTests : IDisposable
{
    private readonly string? _originalPath =
        Environment.GetEnvironmentVariable("EXCELMCP_MAC_VBA_HELPER_PATH");
    private readonly string _directory = Path.Combine(
        Path.GetTempPath(),
        $"excelmcp-helper-client-{Guid.NewGuid():N}");

    [Fact]
    public void Installation_RequiresExactConfiguredHelperFile()
    {
        Environment.SetEnvironmentVariable("EXCELMCP_MAC_VBA_HELPER_PATH", null);
        var missingConfiguration = MacVbaHelperClient.GetInstallation();
        Assert.False(missingConfiguration.IsConfigured);
        Assert.False(missingConfiguration.SourceExists);

        Environment.SetEnvironmentVariable(
            "EXCELMCP_MAC_VBA_HELPER_PATH",
            Path.Combine(_directory, "Imposter.xlam"));
        var wrongName = MacVbaHelperClient.GetInstallation();
        Assert.True(wrongName.IsConfigured);
        Assert.False(wrongName.SourceExists);
        Assert.Contains("must end with ExcelMcpHelper.xlam", wrongName.Status, StringComparison.Ordinal);

        Environment.SetEnvironmentVariable(
            "EXCELMCP_MAC_VBA_HELPER_PATH",
            Path.Combine(_directory, "ExcelMcpHelper.xlam"));
        var missingFile = MacVbaHelperClient.GetInstallation();
        Assert.True(missingFile.IsConfigured);
        Assert.False(missingFile.SourceExists);
        Assert.Contains("does not exist", missingFile.Status, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MutationDispatch_ProbesAndValidatesVersionFirst()
    {
        Directory.CreateDirectory(_directory);
        var helperPath = Path.Combine(_directory, "ExcelMcpHelper.xlam");
        File.WriteAllBytes(helperPath, []);
        Environment.SetEnvironmentVariable("EXCELMCP_MAC_VBA_HELPER_PATH", helperPath);
        var actions = new List<string>();
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            using var envelope = JsonDocument.Parse(input!);
            Assert.Equal(helperPath, envelope.RootElement.GetProperty("helperPath").GetString());
            var requestJson = envelope.RootElement.GetProperty("requestJson").GetString()!;
            var request = MacVbaHelperProtocol.ParseRequest(requestJson);
            actions.Add(request.Action);
            var result = request.Action == "helper.capabilities"
                ? """
                  {"helperVersion":"1.0.0","protocolVersion":1,"supportedActions":["vba.view"],"staticAvailability":{},"engineCapabilities":{},"trustReadiness":{},"provenMethods":{}}
                  """
                : """{"moduleName":"Module1","source":"Option Explicit"}""";
            var response = $$"""
                {"version":1,"requestId":"{{request.RequestId}}","success":true,"result":{{result}},"error":null}
                """;
            var output = JsonSerializer.Serialize(
                new { success = true, errorMessage = "", responseJson = response },
                ServiceProtocol.JsonOptions);
            return Task.FromResult(new MacProcessResult(0, output, ""));
        });
        var client = new MacVbaHelperClient(backend);

        var result = await client.DispatchAsync(
            "/tmp/exact workbook.xlsm",
            "vba.view",
            new { moduleName = "Module1" },
            TimeSpan.FromSeconds(5));

        Assert.Equal(["helper.capabilities", "vba.view"], actions);
        Assert.Equal("Module1", result.GetProperty("moduleName").GetString());
    }

    [Fact]
    public async Task VersionMismatch_PreventsActionDispatch()
    {
        Directory.CreateDirectory(_directory);
        var helperPath = Path.Combine(_directory, "ExcelMcpHelper.xlam");
        File.WriteAllBytes(helperPath, []);
        Environment.SetEnvironmentVariable("EXCELMCP_MAC_VBA_HELPER_PATH", helperPath);
        var calls = 0;
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            calls++;
            using var envelope = JsonDocument.Parse(input!);
            var request = MacVbaHelperProtocol.ParseRequest(
                envelope.RootElement.GetProperty("requestJson").GetString()!);
            const string result =
                """{"helperVersion":"0.9.0","protocolVersion":1,"supportedActions":[],"staticAvailability":{},"engineCapabilities":{},"trustReadiness":{},"provenMethods":{}}""";
            var response = $$"""
                {"version":1,"requestId":"{{request.RequestId}}","success":true,"result":{{result}},"error":null}
                """;
            var output = JsonSerializer.Serialize(
                new { success = true, errorMessage = "", responseJson = response },
                ServiceProtocol.JsonOptions);
            return Task.FromResult(new MacProcessResult(0, output, ""));
        });
        var client = new MacVbaHelperClient(backend);

        var error = await Assert.ThrowsAsync<MacVbaHelperException>(() => client.DispatchAsync(
            "/tmp/exact workbook.xlsm",
            "vba.view",
            new { moduleName = "Module1" },
            TimeSpan.FromSeconds(5)));

        Assert.Equal("helper_version_mismatch", error.Code);
        Assert.Equal(1, calls);
    }

    [Fact]
    public async Task UnsupportedAction_PreventsActionDispatchAfterCapabilityProbe()
    {
        Directory.CreateDirectory(_directory);
        var helperPath = Path.Combine(_directory, "ExcelMcpHelper.xlam");
        File.WriteAllBytes(helperPath, []);
        Environment.SetEnvironmentVariable("EXCELMCP_MAC_VBA_HELPER_PATH", helperPath);
        var calls = 0;
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            calls++;
            using var envelope = JsonDocument.Parse(input!);
            var request = MacVbaHelperProtocol.ParseRequest(
                envelope.RootElement.GetProperty("requestJson").GetString()!);
            const string result =
                """{"helperVersion":"1.0.0","protocolVersion":1,"supportedActions":["vba.view"],"staticAvailability":{},"engineCapabilities":{},"trustReadiness":{},"provenMethods":{}}""";
            var response = $$"""
                {"version":1,"requestId":"{{request.RequestId}}","success":true,"result":{{result}},"error":null}
                """;
            var output = JsonSerializer.Serialize(
                new { success = true, errorMessage = "", responseJson = response },
                ServiceProtocol.JsonOptions);
            return Task.FromResult(new MacProcessResult(0, output, ""));
        });
        var client = new MacVbaHelperClient(backend);

        var error = await Assert.ThrowsAsync<MacVbaHelperException>(() => client.DispatchAsync(
            "/tmp/exact workbook.xlsm",
            "vba.update",
            new { moduleName = "Module1", source = "Option Explicit" },
            TimeSpan.FromSeconds(5)));

        Assert.Equal("helper_action_unavailable", error.Code);
        Assert.Equal(1, calls);
    }

    public void Dispose()
    {
        Environment.SetEnvironmentVariable("EXCELMCP_MAC_VBA_HELPER_PATH", _originalPath);
        if (Directory.Exists(_directory))
        {
            Directory.Delete(_directory, recursive: true);
        }
    }
}
