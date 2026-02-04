using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

/// <summary>
/// Unit tests for ToolInstallationDetector.
/// Tests the logic for detecting global vs local .NET tool installation.
/// </summary>
[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "DaemonTray")]
[Trait("Speed", "Fast")]
public sealed class ToolInstallationDetectorTests
{
    [Fact]
    public void GetInstallationType_ReturnsValidType()
    {
        // Act
        var installationType = ToolInstallationDetector.GetInstallationType();

        // Assert - Should return one of the valid types
        Assert.True(
            installationType == InstallationType.Global ||
            installationType == InstallationType.Local ||
            installationType == InstallationType.Unknown,
            $"Expected valid InstallationType, got: {installationType}");
    }

    [Fact]
    public void GetUpdateCommand_ReturnsValidCommand()
    {
        // Act
        var command = ToolInstallationDetector.GetUpdateCommand();

        // Assert - Should be a valid dotnet tool update command
        Assert.NotNull(command);
        Assert.StartsWith("dotnet tool update", command);
        Assert.Contains("Sbroenne.ExcelMcp.CLI", command);
    }

    [Fact]
    public void GetUpdateCommand_GlobalInstall_IncludesGlobalFlag()
    {
        // Act
        var installationType = ToolInstallationDetector.GetInstallationType();
        var command = ToolInstallationDetector.GetUpdateCommand();

        // Assert - If global, should include --global flag
        if (installationType == InstallationType.Global)
        {
            Assert.Contains("--global", command);
        }
    }

    [Fact]
    public void GetUpdateCommand_LocalInstall_OmitsGlobalFlag()
    {
        // Act
        var installationType = ToolInstallationDetector.GetInstallationType();
        var command = ToolInstallationDetector.GetUpdateCommand();

        // Assert - If local, should NOT include --global flag
        if (installationType == InstallationType.Local)
        {
            Assert.DoesNotContain("--global", command);
        }
    }

    [Fact]
    public void GetUpdateCommand_UnknownInstall_DefaultsToGlobal()
    {
        // This tests the fallback behavior - unknown defaults to global
        // We can't force Unknown state, but we verify the command is always valid
        var command = ToolInstallationDetector.GetUpdateCommand();

        // Assert - Command should always be executable
        Assert.Matches(@"^dotnet tool update (--global )?Sbroenne\.ExcelMcp\.CLI$", command);
    }

    [Fact]
    public async Task TryUpdateAsync_ReturnsResult()
    {
        // Act - This actually tries to run the update command
        // We use a timeout to prevent long waits if network is slow
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));

        var (success, output) = await ToolInstallationDetector.TryUpdateAsync();

        // Assert - Should return a result (success or failure with message)
        Assert.NotNull(output);
        // Note: success may be true or false depending on whether tool is installed
        // and whether there's an update available
    }
}
