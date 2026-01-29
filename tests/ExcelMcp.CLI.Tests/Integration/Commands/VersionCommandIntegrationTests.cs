using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "Version")]
[Trait("Speed", "Fast")]
public sealed class VersionCommandIntegrationTests
{
    [Fact]
    public async Task VersionCommand_Execute_ReturnsSuccess()
    {
        var console = new TestCliConsole();
        var command = new VersionCommand(console);

        // Without --check, VersionReporter.WriteVersion() outputs to Console directly
        // Just verify command returns success exit code
        var exit = await command.ExecuteAsync(null!, new VersionCommand.Settings { Check = false }, CancellationToken.None);

        Assert.Equal(0, exit);
    }

    [Fact]
    public async Task VersionCommand_CheckFlag_QueriesNuGetAsync()
    {
        var console = new TestCliConsole();
        var command = new VersionCommand(console);
        var exit = await command.ExecuteAsync(null!, new VersionCommand.Settings { Check = true }, CancellationToken.None);

        Assert.Equal(0, exit);
        var lastOutput = console.GetLastJson();
        using var json = JsonDocument.Parse(lastOutput);

        // Should have currentVersion
        Assert.True(json.RootElement.TryGetProperty("currentVersion", out var currentVersion));
        Assert.False(string.IsNullOrWhiteSpace(currentVersion.GetString()));

        // Should have updateAvailable (bool)
        Assert.True(json.RootElement.TryGetProperty("updateAvailable", out var updateAvailable));
        // Just check it's a valid boolean - don't care if true/false
        _ = updateAvailable.GetBoolean();

        // Should have latestVersion (might be null if NuGet call fails)
        Assert.True(json.RootElement.TryGetProperty("latestVersion", out _));
    }
}
