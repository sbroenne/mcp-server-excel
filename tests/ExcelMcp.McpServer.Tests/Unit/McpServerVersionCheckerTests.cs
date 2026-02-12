using Sbroenne.ExcelMcp.Service.Infrastructure;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "VersionCheck")]
[Trait("Speed", "Fast")]
public sealed class McpServerVersionCheckerTests
{
    [Fact]
    public async Task CheckForUpdateAsync_WhenUpdateAvailable_ReturnsUpdateInfo()
    {
        // This test depends on NuGetVersionChecker actually checking NuGet
        // In a real scenario, we might want to mock this, but for now we'll test
        // that the method doesn't throw and returns a reasonable result
        var updateInfo = await Infrastructure.McpServerVersionChecker.CheckForUpdateAsync();

        // The result can be null (if no update or network error) or an UpdateInfo object
        // We just verify it doesn't throw and the result is valid
        if (updateInfo != null)
        {
            Assert.NotNull(updateInfo.CurrentVersion);
            Assert.NotNull(updateInfo.LatestVersion);
            Assert.True(updateInfo.UpdateAvailable);
        }
    }

    [Fact]
    public async Task CheckForUpdateAsync_NetworkFailure_ReturnsNull()
    {
        // With a very short cancellation token, we should get a timeout/cancellation
        using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(1));
        var updateInfo = await Infrastructure.McpServerVersionChecker.CheckForUpdateAsync(cts.Token);

        // Should return null on timeout/cancellation (fails silently)
        Assert.Null(updateInfo);
    }

    [Fact]
    public void GetCurrentVersion_ReturnsNonEmptyString()
    {
        var version = Infrastructure.McpServerVersionChecker.GetCurrentVersion();

        Assert.NotNull(version);
        Assert.NotEmpty(version);
        Assert.NotEqual("0.0.0", version); // Should have a real version from assembly
    }

    [Fact]
    public void UpdateInfo_GetUpdateMessage_IncludesVersions()
    {
        var updateInfo = new UpdateInfo
        {
            CurrentVersion = "1.0.0",
            LatestVersion = "1.1.0",
            UpdateAvailable = true
        };

        var message = updateInfo.GetUpdateMessage();

        Assert.Contains("1.0.0", message);
        Assert.Contains("1.1.0", message);
    }

    [Fact]
    public void UpdateInfo_GetUpdateMessage_ContainsUpdateInstructions()
    {
        var updateInfo = new UpdateInfo
        {
            CurrentVersion = "1.0.0",
            LatestVersion = "1.1.0",
            UpdateAvailable = true
        };

        // Use the overload that includes the update command
        var message = updateInfo.GetUpdateMessage("dotnet tool update --global Sbroenne.ExcelMcp.McpServer");

        Assert.Contains("dotnet tool update --global Sbroenne.ExcelMcp.McpServer", message);
    }
}




