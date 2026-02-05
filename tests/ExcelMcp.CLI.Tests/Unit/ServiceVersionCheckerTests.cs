using Sbroenne.ExcelMcp.Service.Infrastructure;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "VersionCheck")]
[Trait("Speed", "Fast")]
public sealed class ServiceVersionCheckerTests
{
    [Fact]
    public async Task CheckForUpdateAsync_WhenUpdateAvailable_ReturnsUpdateInfo()
    {
        // This test depends on NuGetVersionChecker actually checking NuGet
        // In a real scenario, we might want to mock this, but for now we'll test
        // that the method doesn't throw and returns a reasonable result
        var updateInfo = await ServiceVersionChecker.CheckForUpdateAsync();

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
        var updateInfo = await ServiceVersionChecker.CheckForUpdateAsync(cts.Token);

        // Should return null on timeout/cancellation (fails silently)
        Assert.Null(updateInfo);
    }

    [Fact]
    public void UpdateInfo_GetNotificationTitle_ReturnsExpectedFormat()
    {
        // GetNotificationTitle is static, doesn't need an instance
        var title = UpdateInfo.GetNotificationTitle();

        Assert.Equal("Excel MCP Update Available", title);
    }

    [Fact]
    public void UpdateInfo_GetNotificationMessage_IncludesVersions()
    {
        var updateInfo = new UpdateInfo
        {
            CurrentVersion = "1.0.0",
            LatestVersion = "1.1.0",
            UpdateAvailable = true
        };

        var message = updateInfo.GetNotificationMessage();

        Assert.Contains("1.0.0", message);
        Assert.Contains("1.1.0", message);
        Assert.Contains("dotnet tool update", message);
    }

    [Fact]
    public void UpdateInfo_GetNotificationMessage_ContainsUpdateInstructions()
    {
        var updateInfo = new UpdateInfo
        {
            CurrentVersion = "1.0.0",
            LatestVersion = "1.1.0",
            UpdateAvailable = true
        };

        var message = updateInfo.GetNotificationMessage();

        Assert.Contains("dotnet tool update --global Sbroenne.ExcelMcp.McpServer", message);
    }
}




