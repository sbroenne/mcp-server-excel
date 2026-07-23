using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "VersionCheck")]
[Trait("Speed", "Fast")]
public sealed class McpServerVersionCheckerTests
{
    [Fact]
    public async Task CheckForUpdateAsync_WhenCalled_DoesNotThrow()
    {
        // Verify the method doesn't throw — result depends on network/NuGet state
        var latestVersion = await Infrastructure.McpServerVersionChecker.CheckForUpdateAsync();

        // Result can be null (no update or network error) or a version string
        if (latestVersion != null)
        {
            Assert.NotEmpty(latestVersion);
        }
    }

    [Fact]
    public async Task CheckForUpdateAsync_NetworkFailure_ReturnsNull()
    {
        using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(1));
        var latestVersion = await Infrastructure.McpServerVersionChecker.CheckForUpdateAsync(cts.Token);

        Assert.Null(latestVersion);
    }

    [Fact]
    public void GetCurrentVersion_ReturnsNonEmptyString()
    {
        var version = Infrastructure.McpServerVersionChecker.GetCurrentVersion();

        Assert.NotNull(version);
        Assert.NotEmpty(version);
        Assert.NotEqual("0.0.0", version);
    }
}

