using Sbroenne.ExcelMcp.ComInterop.ServiceClient;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

/// <summary>
/// Unit tests for ServiceProtocol version negotiation.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
public class ServiceProtocolTests
{
    [Fact]
    public void Version_ReturnsNonEmptyString()
    {
        var version = ServiceProtocol.Version;
        Assert.False(string.IsNullOrWhiteSpace(version));
    }

    [Fact]
    public void Version_HasMajorMinorPatchFormat()
    {
        var version = ServiceProtocol.Version;
        var parts = version.Split('.');
        Assert.True(parts.Length >= 2, $"Version '{version}' should have at least Major.Minor format");
        Assert.True(int.TryParse(parts[0], out _), $"Major version '{parts[0]}' should be numeric");
        Assert.True(int.TryParse(parts[1], out _), $"Minor version '{parts[1]}' should be numeric");
    }

    [Fact]
    public void Version_DoesNotContainBuildMetadata()
    {
        var version = ServiceProtocol.Version;
        Assert.DoesNotContain("+", version);
    }

    [Fact]
    public void Version_IsConsistentAcrossCalls()
    {
        var version1 = ServiceProtocol.Version;
        var version2 = ServiceProtocol.Version;
        Assert.Equal(version1, version2);
    }
}
