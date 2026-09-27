using Sbroenne.ExcelMcp.Service.Mac;
using System.Text.Json;
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
    [InlineData("analysis.create-scenario")]
    [InlineData("analysis.update-scenario")]
    [InlineData("analysis.show-scenario")]
    [InlineData("analysis.delete-scenario")]
    [InlineData("analysis.create-scenario-summary")]
    public void UnprovenScenarioCommands_RemainExplicitlyGated(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.Contains("real-Excel fixture", capability.UnavailableMessage, StringComparison.Ordinal);
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
    [InlineData("chart.create-from-range")]
    [InlineData("pivottable.create-from-range")]
    public void OfficeAddInCommands_ReportTheirRequiredTier(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.OfficeAddIn, capability.RequiredTier);
        Assert.Contains("Office.js", capability.UnavailableMessage, StringComparison.Ordinal);
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

    [Fact]
    public void GatedCommand_UsesBlockerAsAReasonWithoutDuplicatedPunctuation()
    {
        var capability = MacCommandCapabilities.Get("analysis.goal-seek");

        Assert.Equal(
            "Command 'analysis.goal-seek' is unavailable on macOS: " +
            "No verified macOS backend route exists for this action. " +
            "It remains available on Windows.",
            capability.UnavailableMessage);
    }

    [Fact]
    public void Inventory_ClassifiesEveryPublicActionWithActionLevelEvidence()
    {
        var inventory = MacCommandCapabilities.Inventory;

        Assert.True(inventory.Count > 300);
        Assert.Equal(inventory.Count, inventory.Select(item => item.Command).Distinct(StringComparer.Ordinal).Count());
        Assert.Contains(inventory, item => item.Command == "file.open");
        Assert.Contains(inventory, item => item.Command == "powerquery.refresh");
        Assert.Contains(inventory, item => item.Command == "vba.run");
        Assert.All(inventory, item =>
        {
            Assert.False(string.IsNullOrWhiteSpace(item.Command));
            Assert.False(string.IsNullOrWhiteSpace(item.WindowsSemantics));
            Assert.False(string.IsNullOrWhiteSpace(item.ImplementationStatus));
            Assert.False(string.IsNullOrWhiteSpace(item.Evidence));
            Assert.False(string.IsNullOrWhiteSpace(item.ExcelApiVersion));
            if (!item.IsAvailable)
            {
                Assert.False(string.IsNullOrWhiteSpace(item.Blocker));
            }
        });
    }

    [Fact]
    public void GeneratedMachineReadableInventory_MatchesRuntimeInventory()
    {
        var generated = MacCommandCapabilities.InventoryJson;

        Assert.Contains("\"command\": \"sheet.create\"", generated, StringComparison.Ordinal);
        Assert.Contains("\"tier\": \"Native\"", generated, StringComparison.Ordinal);
        Assert.Contains("\"command\": \"powerquery.update\"", generated, StringComparison.Ordinal);
        Assert.Equal(MacCommandCapabilities.Inventory.Count, CountOccurrences(generated, "\"command\":"));
    }

    [Fact]
    public void GeneratedRepositoryInventories_AreCurrent()
    {
        var root = FindRepository();
        using var expected = JsonDocument.Parse(MacCommandCapabilities.InventoryJson);
        using var committed = JsonDocument.Parse(File.ReadAllText(
            Path.Combine(root, "docs", "generated", "macos-action-inventory.json")));
        Assert.True(JsonElement.DeepEquals(expected.RootElement, committed.RootElement));

        var markdown = File.ReadAllText(
            Path.Combine(root, "docs", "MACOS-ACTION-INVENTORY.md"));
        Assert.Contains("Do not edit this table directly", markdown, StringComparison.Ordinal);
        Assert.Equal(
            MacCommandCapabilities.Inventory.Count,
            markdown.Split('\n').Count(line => line.StartsWith("| `", StringComparison.Ordinal)));
    }

    private static int CountOccurrences(string value, string token)
    {
        var count = 0;
        var index = 0;
        while ((index = value.IndexOf(token, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += token.Length;
        }

        return count;
    }

    private static string FindRepository()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null
               && !File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln")))
        {
            directory = directory.Parent;
        }

        return directory?.FullName
            ?? throw new InvalidOperationException("Repository root not found.");
    }
}
