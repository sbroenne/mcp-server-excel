using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

/// <summary>
/// Pure probe decision and caching tests, not substitutes for older Excel integration coverage.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelContext")]
[Trait("RequiresExcel", "false")]
[System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201",
    Justification = "Synthetic HRESULTs exercise pure probe policy, not Excel COM behavior.")]
public sealed class ExcelCapabilitiesTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SupportsFormula2_CachesSuccessfulDecision(bool supported)
    {
        int calls = 0;
        var capabilities = new ExcelCapabilities(() => { calls++; return supported; });

        Assert.Equal(supported, capabilities.SupportsFormula2);
        Assert.Equal(supported, capabilities.SupportsFormula2);
        Assert.Equal(1, calls);
    }

    [Theory]
    [InlineData(unchecked((int)0x80020003))]
    [InlineData(unchecked((int)0x80020006))]
    [InlineData(unchecked((int)0x80004001))]
    [InlineData(unchecked((int)0x800A03EC))]
    public void ProbeFormula2_UnavailableGetter_SelectsLegacy(int hresult)
    {
        Assert.False(ExcelCapabilities.ProbeFormula2(
            () => null,
            () => throw new COMException("Unavailable getter", hresult)));
    }

    [Fact]
    public void ProbeFormula2_ReadableGetter_SelectsModern()
    {
        Assert.True(ExcelCapabilities.ProbeFormula2(() => null, () => null));
    }

    [Fact]
    public void ProbeFormula2_UnreadableLegacyCell_PropagatesFailure()
    {
        var error = new COMException("Unreadable cell", unchecked((int)0x800A03EC));
        bool modernRead = false;
        Assert.Same(error, Assert.Throws<COMException>(() => ExcelCapabilities.ProbeFormula2(
            () => throw error,
            () => { modernRead = true; return null; })));
        Assert.False(modernRead);
    }

    [Theory]
    [InlineData(unchecked((int)0x80010001))]
    [InlineData(unchecked((int)0x80010108))]
    [InlineData(unchecked((int)0x8007000E))]
    public void SupportsFormula2_UnexpectedProbeError_PropagatesWithoutCaching(int hresult)
    {
        int calls = 0;
        var error = new COMException("Unexpected failure", hresult);
        var capabilities = new ExcelCapabilities(() =>
        {
            calls++;
            return ExcelCapabilities.ProbeFormula2(() => null, () => throw error);
        });

        Assert.Same(error, Assert.Throws<COMException>(() => capabilities.SupportsFormula2));
        Assert.Same(error, Assert.Throws<COMException>(() => capabilities.SupportsFormula2));
        Assert.Equal(2, calls);
    }
}
