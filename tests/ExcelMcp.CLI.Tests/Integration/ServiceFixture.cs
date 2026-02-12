using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Fixture that ensures the ExcelMCP service is running before tests execute.
/// The service auto-starts when excelcli connects, but having it pre-started
/// avoids startup timeouts in parallel test execution.
/// </summary>
public sealed class ServiceFixture : IAsyncLifetime
{
    public async Task InitializeAsync()
    {
        // Ensure service is running (will start it if needed)
        var running = await ServiceManager.EnsureServiceRunningAsync();
        if (!running)
        {
            throw new InvalidOperationException(
                "Failed to start ExcelMCP service for integration tests. " +
                "Ensure excelcli.exe is available in the build output.");
        }
    }

    public Task DisposeAsync()
    {
        // Don't stop the service — other test classes may still need it,
        // and it has an idle timeout for auto-shutdown.
        return Task.CompletedTask;
    }
}

/// <summary>
/// Collection definition for tests that require the ExcelMCP service.
/// Apply [Collection("Service")] to test classes that call excelcli commands.
/// </summary>
[CollectionDefinition("Service")]
public sealed class ServiceTestGroup : ICollectionFixture<ServiceFixture>;
