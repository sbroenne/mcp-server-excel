using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "VersionCheck")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class McpServerVersionCheckerTests
{
    [Theory]
    [InlineData("1.2.4", "1.2.4")]
    [InlineData("1.2.3", null)]
    [InlineData("1.2.2", null)]
    public async Task CheckForUpdateAsync_ReturnsOnlyNewerVersions(string latestVersion, string? expected)
    {
        var version = await Infrastructure.McpServerVersionChecker.CheckForUpdateAsync(
            () => "1.2.3",
            _ => Task.FromResult<string?>(latestVersion));

        Assert.Equal(expected, version);
    }

    [Fact]
    public async Task CheckForUpdateAsync_WhenLookupFails_ReturnsNull()
    {
        var version = await Infrastructure.McpServerVersionChecker.CheckForUpdateAsync(
            () => "1.2.3",
            _ => Task.FromException<string?>(new HttpRequestException("Synthetic lookup failure.")));

        Assert.Null(version);
    }

    [Fact]
    public async Task CheckForUpdateAsync_WhenCurrentVersionIsUnavailable_ReturnsNullWithoutLookup()
    {
        var lookupCalled = false;
        var version = await Infrastructure.McpServerVersionChecker.CheckForUpdateAsync(
            () => throw new InvalidOperationException("Synthetic version metadata failure."),
            _ =>
            {
                lookupCalled = true;
                return Task.FromResult<string?>("1.2.4");
            });

        Assert.Null(version);
        Assert.False(lookupCalled);
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
