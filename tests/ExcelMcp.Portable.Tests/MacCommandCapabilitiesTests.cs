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
    [InlineData("rangeedit.insert-cells")]
    [InlineData("rangeedit.delete-cells")]
    [InlineData("rangeedit.insert-rows")]
    [InlineData("rangeedit.delete-rows")]
    [InlineData("rangeedit.insert-columns")]
    [InlineData("rangeedit.delete-columns")]
    public void NativeCommands_AreAvailableWithoutOptionalHelpers(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.True(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
    }

    [Theory]
    [InlineData("diag.ping")]
    [InlineData("diag.echo")]
    [InlineData("diag.validate-params")]
    public void DiagnosticCommands_ArePlatformIndependent(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.True(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.Equal("Implemented", capability.ImplementationStatus);
        Assert.Equal("No Excel API required.", capability.ExcelApiVersion);
    }

    [Fact]
    public void PythonInExcel_SetFormulaIsNative()
    {
        var capability = MacCommandCapabilities.Get("pythoninexcel.set-formula");

        Assert.True(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.Equal("Implemented", capability.ImplementationStatus);
        Assert.Contains("CLI and MCP", capability.Evidence, StringComparison.Ordinal);
        Assert.Contains("16.113.2", capability.ExcelApiVersion, StringComparison.Ordinal);
        Assert.Empty(capability.UnavailableMessage);
    }

    [Fact]
    public void PythonInExcel_GetResultIsBlockedAfterPublicMcpFailure()
    {
        var capability = MacCommandCapabilities.Get("pythoninexcel.get-result");

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.Equal("Blocked", capability.ImplementationStatus);
        Assert.Contains("MCP", capability.Evidence, StringComparison.Ordinal);
        Assert.Contains("Message not understood", capability.Evidence, StringComparison.Ordinal);
        Assert.Contains("16.113.2", capability.ExcelApiVersion, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("sheet.copy")]
    [InlineData("sheet.move")]
    public void SheetReordering_UsesDisabledOfficeAddInCandidate(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.OfficeAddIn, capability.RequiredTier);
        Assert.Equal("Partial", capability.ImplementationStatus);
    }

    [Theory]
    [InlineData("range.copy")]
    [InlineData("range.copy-values")]
    [InlineData("range.copy-formulas")]
    [InlineData("range.get-info")]
    [InlineData("range.set-number-formats")]
    [InlineData("rangeformat.auto-fit-columns")]
    [InlineData("rangeformat.auto-fit-rows")]
    [InlineData("rangeformat.merge-cells")]
    [InlineData("rangeformat.unmerge-cells")]
    [InlineData("rangelink.set-cell-lock")]
    [InlineData("rangelink.get-cell-lock")]
    public void RangeExpansion_ProvenCommandsAreNative(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.True(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.Equal("Implemented", capability.ImplementationStatus);
        Assert.Contains("CLI and MCP", capability.Evidence, StringComparison.Ordinal);
        Assert.Contains("16.113.1", capability.ExcelApiVersion, StringComparison.Ordinal);
        Assert.Empty(capability.UnavailableMessage);
    }

    [Theory]
    [InlineData("range.get-current-region")]
    [InlineData("range.get-used-range")]
    [InlineData("rangeformat.get-merge-info")]
    public void DisprovenRangeDiscoveryCommandsRemainGated(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.Equal("Blocked", capability.ImplementationStatus);
    }

    [Fact]
    public void UsedRange_RemainsBlockedAfterPopulatedSheetRoundTripFailure()
    {
        var capability = MacCommandCapabilities.Get("range.get-used-range");

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.Equal("Blocked", capability.ImplementationStatus);
        Assert.Contains("populated", capability.Evidence, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("missing object", capability.Evidence, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("-50", capability.Evidence, StringComparison.Ordinal);
    }

    [Fact]
    public void MergeInfo_RemainsBlockedWithoutMergeAreaReadback()
    {
        var capability = MacCommandCapabilities.Get("rangeformat.get-merge-info");

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.Equal("Blocked", capability.ImplementationStatus);
        Assert.Contains("merge area", capability.Evidence, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("missing object", capability.Evidence, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("-50", capability.Evidence, StringComparison.Ordinal);
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
        Assert.Equal("Implemented", capability.ImplementationStatus);
        Assert.Contains("CLI and MCP", capability.Evidence, StringComparison.Ordinal);
        Assert.Contains("16.113.1", capability.ExcelApiVersion, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("analysis.list-scenarios")]
    [InlineData("analysis.update-scenario")]
    [InlineData("analysis.delete-scenario")]
    [InlineData("analysis.create-scenario-summary")]
    public void UnverifiedNativeScenarioActions_RemainGated(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.Equal("Partial", capability.ImplementationStatus);
        Assert.Contains("dictionary", capability.Evidence, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("runtime parity", capability.Blocker, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("analysis.create-scenario")]
    [InlineData("analysis.show-scenario")]
    public void ScenarioMutationsWithoutSupportedApis_ReportMacLimitation(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Unsupported, capability.RequiredTier);
        Assert.Equal("Blocked", capability.ImplementationStatus);
        Assert.Equal("MacLimitation", capability.PlannedTier);
        Assert.Contains("no Scenario API", capability.Evidence, StringComparison.Ordinal);
        Assert.Contains("does not ship a VBA helper", capability.Blocker, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("connection.list")]
    [InlineData("connection.view")]
    [InlineData("connection.create")]
    [InlineData("connection.refresh")]
    [InlineData("connection.get-refresh-status")]
    [InlineData("connection.cancel-refresh")]
    [InlineData("connection.delete")]
    [InlineData("connection.load-to")]
    [InlineData("connection.get-properties")]
    [InlineData("connection.set-properties")]
    [InlineData("connection.test")]
    public void ConnectionActions_RequireNewTierOrLimitationEvidence(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Unsupported, capability.RequiredTier);
        Assert.Equal("MacLimitation", capability.PlannedTier);
        Assert.Equal("Blocked", capability.ImplementationStatus);
        Assert.Contains("no workbook collection", capability.Evidence, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("public contract", capability.Blocker, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("querytable.create-text")]
    [InlineData("querytable.create-web")]
    public void QueryTableCreation_RequiresNewTierOrLimitationEvidence(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Unsupported, capability.RequiredTier);
        Assert.Equal("MacLimitation", capability.PlannedTier);
        Assert.Equal("Blocked", capability.ImplementationStatus);
        Assert.Contains("no construction command", capability.Evidence, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("faithful creation route", capability.Blocker, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("querytable.view")]
    [InlineData("querytable.set-properties")]
    public void QueryTableIncompleteContracts_RequireNewTierOrLimitationEvidence(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Unsupported, capability.RequiredTier);
        Assert.Equal("MacLimitation", capability.PlannedTier);
        Assert.Equal("Blocked", capability.ImplementationStatus);
        Assert.Contains("omits", capability.Evidence, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("public", capability.Blocker, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("drawing.add-sparkline")]
    [InlineData("drawing.add-shape")]
    [InlineData("slicer.list-slicers")]
    [InlineData("slicer.set-table-slicer-selection")]
    public void SpecializedOfficeJsCommandsWithoutFaithfulHandlers_ReportLimitations(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Unsupported, capability.RequiredTier);
        Assert.Equal("Blocked", capability.ImplementationStatus);
        Assert.Equal("MacLimitation", capability.PlannedTier);
    }

    [Fact]
    public void Screenshot_ReportsOptionalNativeHelperTier()
    {
        var capability = MacCommandCapabilities.Get("screenshot.capture");

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.OptionalNativeHelper, capability.RequiredTier);
        Assert.Contains("Screen Recording", capability.UnavailableMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void UnprovenQueryTableCandidateReportsNativeTier()
    {
        var capability = MacCommandCapabilities.Get("querytable.list");

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

    [Theory]
    [InlineData("table.create")]
    [InlineData("chart.create-from-range")]
    [InlineData("chartconfig.get-plot-options")]
    [InlineData("pivottable.create-from-range")]
    [InlineData("pivottablefield.set-field-filter")]
    [InlineData("pivottablecalc.get-data")]
    [InlineData("pivottablecalc.set-grand-totals")]
    [InlineData("slicer.create-table-slicer")]
    public void OfficeAddInCandidate_IsRoutableOnlyWhenExplicitlyEnabled(string command)
    {
        var capability = MacCommandCapabilities.Get(
            command,
            officeCandidateEnabled: true);

        Assert.True(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.OfficeAddIn, capability.RequiredTier);
        Assert.Empty(capability.UnavailableMessage);
    }

    [Theory]
    [InlineData("powerquery.list")]
    [InlineData("powerquery.view")]
    [InlineData("powerquery.get-load-config")]
    [InlineData("powerquery.update")]
    [InlineData("powerquery.refresh")]
    [InlineData("powerquery.refresh-all")]
    [InlineData("powerquery.create")]
    [InlineData("powerquery.rename")]
    [InlineData("powerquery.delete")]
    [InlineData("powerquery.load-to")]
    [InlineData("powerquery.unload")]
    [InlineData("powerquery.evaluate")]
    public void PowerQueryCommands_ReportSupportedApiLimitation(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Unsupported, capability.RequiredTier);
        Assert.Equal("Blocked", capability.ImplementationStatus);
        Assert.Equal("MacLimitation", capability.PlannedTier);
        Assert.Contains("no Workbook.Queries", capability.Evidence, StringComparison.Ordinal);
        Assert.Contains("does not ship a VBA helper", capability.Blocker, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("vba.list")]
    [InlineData("vba.view")]
    [InlineData("vba.import")]
    [InlineData("vba.update")]
    [InlineData("vba.delete")]
    [InlineData("vba.run")]
    public void VbaCommands_ReportSupportedApiLimitation(string command)
    {
        var capability = MacCommandCapabilities.Get(command);

        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Unsupported, capability.RequiredTier);
        Assert.Equal("Blocked", capability.ImplementationStatus);
        Assert.Equal("MacLimitation", capability.PlannedTier);
        Assert.Contains("does not ship a VBA helper", capability.Blocker, StringComparison.Ordinal);
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
        var capability = MacCommandCapabilities.Get("analysis.list-scenarios");

        Assert.Equal(
            "Command 'analysis.list-scenarios' is unavailable on macOS: " +
            "native API presence is not runtime parity; exact returned metadata must pass prompt-free CLI and MCP Excel tests. " +
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
                Assert.False(string.IsNullOrWhiteSpace(item.PlannedTier));
                Assert.False(string.IsNullOrWhiteSpace(item.AcceptanceFixture));
                Assert.False(string.IsNullOrWhiteSpace(item.AcceptanceCommand));
                Assert.False(string.IsNullOrWhiteSpace(item.RecoveryRule));
                Assert.False(string.IsNullOrWhiteSpace(item.EvidenceCriteria));
            }
        });
    }

    [Fact]
    public void Inventory_HasNoUnreviewedActions()
    {
        var unreviewed = MacCommandCapabilities.Inventory
            .Where(item => item.ImplementationStatus == "NotTested")
            .Select(item => item.Command)
            .ToArray();

        Assert.Empty(unreviewed);
    }

    [Fact]
    public void OfficeAddInInventory_DistinguishesImplementedCandidatesFromApiLimitations()
    {
        var candidates = MacOfficeActionCatalog.All
            .Select(item => item.Command)
            .ToHashSet(StringComparer.Ordinal);

        Assert.NotEmpty(candidates);
        Assert.All(
            MacCommandCapabilities.Inventory.Where(item => candidates.Contains(item.Command)),
            item =>
            {
                Assert.Equal(MacCapabilityTier.OfficeAddIn, item.RequiredTier);
                Assert.Equal("Partial", item.ImplementationStatus);
                Assert.False(item.IsAvailable);
                Assert.Contains("activate the task-pane", item.Blocker, StringComparison.OrdinalIgnoreCase);
            });
        Assert.All(
            MacCommandCapabilities.Inventory.Where(item =>
                item.PlannedTier == "MacLimitation"
                && item.Command.StartsWith("chart.", StringComparison.Ordinal)),
            item =>
            {
                Assert.Equal(MacCapabilityTier.Unsupported, item.RequiredTier);
                Assert.Equal("Blocked", item.ImplementationStatus);
                Assert.DoesNotContain(item.Command, candidates);
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
    public void GeneratedMachineReadableInventory_ClassifiesEveryGatedActionForExecution()
    {
        using var inventory = JsonDocument.Parse(MacCommandCapabilities.InventoryJson);

        foreach (var action in inventory.RootElement.EnumerateArray()
                     .Where(item => !item.GetProperty("isAvailable").GetBoolean()))
        {
            Assert.False(string.IsNullOrWhiteSpace(
                action.GetProperty("plannedTier").GetString()));
            Assert.False(string.IsNullOrWhiteSpace(
                action.GetProperty("acceptanceFixture").GetString()));
            Assert.False(string.IsNullOrWhiteSpace(
                action.GetProperty("acceptanceCommand").GetString()));
            Assert.False(string.IsNullOrWhiteSpace(
                action.GetProperty("recoveryRule").GetString()));
            Assert.False(string.IsNullOrWhiteSpace(
                action.GetProperty("evidenceCriteria").GetString()));
        }
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
        Assert.Contains("**Planned tiers for gated actions:**", markdown, StringComparison.Ordinal);
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
