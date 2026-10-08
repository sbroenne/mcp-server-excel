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
    [InlineData(unchecked((int)0x80020003), false)]
    [InlineData(unchecked((int)0x80020006), false)]
    [InlineData(unchecked((int)0x80004001), false)]
    [InlineData(unchecked((int)0x80020003), true)]
    [InlineData(unchecked((int)0x80020006), true)]
    [InlineData(unchecked((int)0x80004001), true)]
    public void AutoSave_UnavailableMember_ReadsFalseAndAllowsDisable(int hresult, bool writing)
    {
        if (writing)
            ExcelCapabilities.DisableAutoSave(() => throw new COMException("Unsupported AutoSave", hresult));
        else
            Assert.False(ExcelCapabilities.ReadAutoSave(() => throw new COMException("Unsupported AutoSave", hresult)));
    }

    [Theory]
    [InlineData(unchecked((int)0x800A03EC))]
    [InlineData(unchecked((int)0x80010001))]
    [InlineData(unchecked((int)0x80010108))]
    [InlineData(unchecked((int)0x80070005))]
    public void AutoSave_UnexpectedReadOrWriteFailure_Propagates(int hresult)
    {
        var error = new COMException("AutoSave failed", hresult);
        Assert.Same(error, Assert.Throws<COMException>(() => ExcelCapabilities.ReadAutoSave(() => throw error)));
        Assert.Same(error, Assert.Throws<COMException>(() => ExcelCapabilities.DisableAutoSave(() => throw error)));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void AutoSave_SupportedMember_ReturnsLiveValueAndDisables(bool enabled)
    {
        Assert.Equal(enabled, ExcelCapabilities.ReadAutoSave(() => enabled));
        enabled = !enabled;
        Assert.Equal(enabled, ExcelCapabilities.ReadAutoSave(() => enabled));
        ExcelCapabilities.DisableAutoSave(() => enabled = false);
        Assert.False(enabled);
    }

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
