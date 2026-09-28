using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacCommandCapabilitiesTests
{
    [Theory]
    [InlineData("sheet.create")]
    [InlineData("sheet.set-visibility")]
    [InlineData("sheet.get-visibility")]
    [InlineData("sheet.show")]
    [InlineData("sheet.hide")]
    [InlineData("sheet.very-hide")]
    [InlineData("sheet.set-tab-color")]
    [InlineData("sheet.get-tab-color")]
    [InlineData("sheet.clear-tab-color")]
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
    [InlineData("pythoninexcel.set-formula")]
    [InlineData("pythoninexcel.get-result")]
    public void PythonInExcel_RemainsGatedWithoutPersistentFormula2RoundTrip(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Unsupported, capability.RequiredTier);
    }

    [Theory]
    [InlineData("sheet.copy")]
    [InlineData("sheet.move")]
    public void SheetReordering_RemainsGatedWithoutProvenAppleEventsParity(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Unsupported, capability.RequiredTier);
    }

    [Theory]
    [InlineData("range.copy")]
    [InlineData("range.copy-values")]
    [InlineData("range.copy-formulas")]
    [InlineData("range.get-current-region")]
    [InlineData("range.get-used-range")]
    [InlineData("range.get-info")]
    [InlineData("range.set-number-formats")]
    [InlineData("rangeformat.auto-fit-columns")]
    [InlineData("rangeformat.auto-fit-rows")]
    [InlineData("rangeformat.merge-cells")]
    [InlineData("rangeformat.unmerge-cells")]
    [InlineData("rangeformat.get-merge-info")]
    [InlineData("rangelink.set-cell-lock")]
    [InlineData("rangelink.get-cell-lock")]
    public void RangeExpansion_RemainsGatedUntilBothEntryPointsPass(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Unsupported, capability.RequiredTier);
        Assert.Contains("not completed real CLI and MCP acceptance", capability.UnavailableMessage);
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
    [InlineData("powerquery.get-load-config")]
    [InlineData("powerquery.update")]
    public void PowerQueryReadCommands_AreAvailableThroughSecurePackageTier(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.True(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.PowerQueryPackage, capability.RequiredTier);
        Assert.Empty(capability.UnavailableMessage);
    }

    [Theory]
    [InlineData("powerquery.refresh")]
    [InlineData("powerquery.refresh-all")]
    public void PowerQueryRefresh_StaysGatedWithoutRealExcelCompletionEvidence(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.PowerQueryPackage, capability.RequiredTier);
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
