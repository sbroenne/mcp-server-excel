using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacCommandCapabilitiesTests
{
    [Theory]
    [InlineData("sheet.create")]
    [InlineData("range.get-number-formats")]
    [InlineData("range.set-number-format")]
    [InlineData("rangeformat.set-column-width")]
    [InlineData("rangeformat.set-row-height")]
    public void NativeCommands_AreAvailableWithoutOptionalHelpers(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.True(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
    }

    [Theory]
    [InlineData("table.create")]
    [InlineData("chart.create")]
    [InlineData("pivottable.create")]
    public void OfficeAddInCommands_ReportTheirRequiredTier(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.OfficeAddIn, capability.RequiredTier);
        Assert.Contains("Office.js", capability.UnavailableMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void VbaRun_RequiresMacroHelperWithoutProjectModelTrust()
    {
        var capability = MacCommandCapabilities.Get("vba.run");

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.MacroHelper, capability.RequiredTier);
        Assert.DoesNotContain("project object model", capability.UnavailableMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("powerquery.list")]
    [InlineData("powerquery.view")]
    public void PowerQueryReadCommands_AreAvailableThroughSecurePackageTier(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.True(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.PowerQueryPackage, capability.RequiredTier);
        Assert.Empty(capability.UnavailableMessage);
    }

    [Fact]
    public void PowerQueryMutation_StaysGatedUntilWorkbookTransactionIsImplemented()
    {
        var capability = MacCommandCapabilities.Get("powerquery.update");

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.PowerQueryPackage, capability.RequiredTier);
        Assert.Contains("saved-package", capability.UnavailableMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("vba.list")]
    [InlineData("vba.view")]
    [InlineData("vba.import")]
    [InlineData("vba.update")]
    [InlineData("vba.delete")]
    public void VbaSourceCommands_RequireProjectModelTrust(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.VbaProjectModel, capability.RequiredTier);
        Assert.Contains("project object model", capability.UnavailableMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void UnknownCommand_RemainsExplicitlyUnsupported()
    {
        var capability = MacCommandCapabilities.Get("datamodel.create-table");

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Unsupported, capability.RequiredTier);
        Assert.Contains("not supported", capability.UnavailableMessage, StringComparison.OrdinalIgnoreCase);
    }
}
