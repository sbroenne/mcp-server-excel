using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Fixture that starts an in-process ExcelMCP service for CLI integration tests.
/// Uses the CLI pipe name so CLI commands can connect to it.
/// </summary>
public sealed class ServiceFixture : IAsyncLifetime, IDisposable
{
    private ExcelMcpService? _service;
    private string? _pipeName;
    private string? _previousPipeNameOverride;

    public async Task InitializeAsync()
    {
        _pipeName = $"excelmcp-cli-test-{Guid.NewGuid():N}";
        _previousPipeNameOverride = Environment.GetEnvironmentVariable("EXCELMCP_CLI_PIPE");
        Environment.SetEnvironmentVariable("EXCELMCP_CLI_PIPE", _pipeName);

        _service = new ExcelMcpService();
        _ = Task.Run(() => _service.RunAsync(_pipeName));

        // Wait for pipe server to be ready
        for (int i = 0; i < 20; i++)
        {
            await Task.Delay(100);
            using var client = new ServiceClient(_pipeName, connectTimeout: TimeSpan.FromSeconds(1));
            if (await client.PingAsync())
            {
                return;
            }
        }

        throw new InvalidOperationException("ExcelMCP service did not start within timeout.");
    }

    public Task DisposeAsync()
    {
        Dispose();
        return Task.CompletedTask;
    }

    public void Dispose()
    {
        _service?.RequestShutdown();
        _service?.Dispose();
        _service = null;
        Environment.SetEnvironmentVariable("EXCELMCP_CLI_PIPE", _previousPipeNameOverride);
        _previousPipeNameOverride = null;
        _pipeName = null;
    }
}

/// <summary>
/// Collection definition for tests that require the ExcelMCP service.
/// Apply [Collection("Service")] to test classes that call excelcli commands.
/// </summary>
[CollectionDefinition("Service")]
public sealed class ServiceTestGroup : ICollectionFixture<ServiceFixture>;
