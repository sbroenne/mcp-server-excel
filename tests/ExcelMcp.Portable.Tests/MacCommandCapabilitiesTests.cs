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
    [InlineData("analysis.goal-seek")]
    [InlineData("analysis.create-data-table")]
    public void ProvenWhatIfAnalysisCommands_AreNative(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.True(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.Empty(capability.UnavailableMessage);
    }

    [Theory]
    [InlineData("analysis.list-scenarios")]
    [InlineData("analysis.update-scenario")]
    [InlineData("analysis.delete-scenario")]
    [InlineData("analysis.create-scenario-summary")]
    public void DictionaryBackedScenarioCommands_RemainGatedUntilRealExcelEvidence(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.Contains("real-Excel fixture", capability.UnavailableMessage, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("analysis.create-scenario")]
    [InlineData("analysis.show-scenario")]
    public void ScenarioCommandsMissingFromNativeDictionary_ReportMacroHelperTier(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.MacroHelper, capability.RequiredTier);
        Assert.Contains("VBA helper", capability.UnavailableMessage, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("drawing.add-sparkline")]
    [InlineData("drawing.add-shape")]
    [InlineData("slicer.list-slicers")]
    [InlineData("slicer.set-table-slicer-selection")]
    public void SpecializedOfficeJsCommands_ReportAddInTier(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.OfficeAddIn, capability.RequiredTier);
    }

    [Fact]
    public void Screenshot_ReportsOptionalNativeHelperTier()
    {
        var capability = MacCommandCapabilities.Get("screenshot.capture");

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.OptionalNativeHelper, capability.RequiredTier);
        Assert.Contains("Screen Recording", capability.UnavailableMessage, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("connection.list")]
    [InlineData("querytable.list")]
    [InlineData("pythoninexcel.set-formula")]
    public void UnprovenAppleEventCandidates_ReportNativeTier(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.Contains("real-Excel fixture", capability.UnavailableMessage, StringComparison.Ordinal);
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
    public void OfficeAddInCandidate_IsRoutableOnlyWhenExplicitlyEnabled()
    {
        var capability = MacCommandCapabilities.Get(
            "table.create",
            officeCandidateEnabled: true);

        Assert.True(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.OfficeAddIn, capability.RequiredTier);
        Assert.Empty(capability.UnavailableMessage);
    }

    [Fact]
    public void VbaRun_RemainsGatedWithoutRepositoryFixtureEvidence()
    {
        var capability = MacCommandCapabilities.Get(
            "vba.run",
            new MacVbaPreflightResult(
                MacMacroExecutionAvailability.Available,
                MacVbaProjectModelAccess.Disabled));

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.MacroHelper, capability.RequiredTier);
        Assert.DoesNotContain("project object model", capability.UnavailableMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("repository-owned", capability.UnavailableMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("unattended", capability.UnavailableMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("Disabled", "disabled")]
    [InlineData("UserApprovalRequired", "approval")]
    [InlineData("Unknown", "could not")]
    public void VbaRun_ReportsNonPromptingMacroPreflight(
        string availabilityName,
        string expectedMessage)
    {
        var availability = Enum.Parse<MacMacroExecutionAvailability>(availabilityName);
        var capability = MacCommandCapabilities.Get(
            "vba.run",
            new MacVbaPreflightResult(availability, MacVbaProjectModelAccess.Disabled));

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.MacroHelper, capability.RequiredTier);
        Assert.Contains(expectedMessage, capability.UnavailableMessage, StringComparison.OrdinalIgnoreCase);
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
    public void VbaSourceCommands_ReportMissingTrustAndScriptingRoute(string command)
    {
        var capability = MacCommandCapabilities.Get(
            command,
            new MacVbaPreflightResult(
                MacMacroExecutionAvailability.Available,
                MacVbaProjectModelAccess.Disabled));

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.VbaProjectModel, capability.RequiredTier);
        Assert.Contains("project object model", capability.UnavailableMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("disabled", capability.UnavailableMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("scripting", capability.UnavailableMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void VbaSourceCommands_StayGatedWhenProjectModelTrustIsEnabled()
    {
        var capability = MacCommandCapabilities.Get(
            "vba.list",
            new MacVbaPreflightResult(
                MacMacroExecutionAvailability.Available,
                MacVbaProjectModelAccess.Enabled));

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.VbaProjectModel, capability.RequiredTier);
        Assert.Contains("enabled", capability.UnavailableMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("scripting", capability.UnavailableMessage, StringComparison.OrdinalIgnoreCase);
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
