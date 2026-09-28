using Sbroenne.ExcelMcp.Generated;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal enum MacCapabilityTier
{
    Native,
    OfficeAddIn,
    MacroHelper,
    VbaProjectModel,
    OptionalNativeHelper,
    Unsupported
}

internal sealed record MacCommandCapability(
    string Command,
    bool IsAvailable,
    MacCapabilityTier RequiredTier,
    string UnavailableMessage,
    string WindowsSemantics,
    IReadOnlyList<string> WindowsVariants,
    string ImplementationStatus,
    string Evidence,
    string ExcelApiVersion,
    string Blocker);

internal static class MacCommandCapabilities
{
    private static readonly IReadOnlyList<MacCommandCapability> All =
        MacActionInventory.Actions.Select(ToCapability).ToArray();

    private static readonly Dictionary<string, MacCommandCapability> ByCommand =
        All.ToDictionary(item => item.Command, StringComparer.Ordinal);

    public static IReadOnlyList<MacCommandCapability> Inventory => All;

    public static string InventoryJson => MacActionInventory.Json;

    public static MacCommandCapability Get(
        string command,
        MacVbaPreflightResult? vbaPreflight = null,
        bool officeCandidateEnabled = false)
    {
        "sheet.list",
        "sheet.create",
        "sheet.rename",
        "sheet.delete",
        "range.get-values",
        "range.set-values",
        "range.get-formulas",
        "range.set-formulas",
        "range.clear-all",
        "range.clear-contents",
        "range.clear-formats",
        "range.get-number-formats",
        "range.set-number-format",
        "rangeformat.set-column-width",
        "rangeformat.set-row-height",
        "calculation.calculate"
    };

    private static readonly HashSet<string> PendingNativeRangeCommands = new(StringComparer.Ordinal)
    {
        "range.copy",
        "range.copy-values",
        "range.copy-formulas",
        "range.get-current-region",
        "range.get-used-range",
        "range.get-info",
        "range.set-number-formats",
        "rangeformat.auto-fit-columns",
        "rangeformat.auto-fit-rows",
        "rangeformat.merge-cells",
        "rangeformat.unmerge-cells",
        "rangeformat.get-merge-info",
        "rangelink.set-cell-lock",
        "rangelink.get-cell-lock"
    };

    private static readonly HashSet<string> OfficeAddInCategories = new(StringComparer.Ordinal)
    {
        "table",
        "tablecolumn",
        "chart",
        "chartconfig",
        "pivottable",
        "pivottablefield",
        "pivottablecalc",
        "conditionalformat"
    };

    public static MacCommandCapability Get(string command)
    {
        if (NativeCommands.Contains(command))
        {
            if (officeCandidateEnabled && MacOfficeActionCatalog.TryGet(command, out _))
            {
                return capability with
                {
                    IsAvailable = true,
                    RequiredTier = MacCapabilityTier.OfficeAddIn,
                    UnavailableMessage = string.Empty
                };
            }
            if (!capability.IsAvailable && command.StartsWith("vba.", StringComparison.Ordinal))
            {
                vbaPreflight ??= MacVbaPreflight.Check();
                var readiness = command == "vba.run"
                    ? MacVbaPreflight.DescribeMacroExecution(vbaPreflight.MacroExecution)
                    : MacVbaPreflight.DescribeProjectModel(vbaPreflight.ProjectModelAccess);
                var evidence = command == "vba.run"
                    ? "A repository-owned fixture has not yet proven unattended workbook-qualified execution through CLI and MCP."
                    : "Apple Events scripting exposes no project-model route; the optional helper requires separate execution evidence.";
                return capability with
                {
                    UnavailableMessage = $"{capability.UnavailableMessage} Preflight reports that {readiness}. {evidence}"
                };
            }
            return capability;
        }

        if (PendingNativeRangeCommands.Contains(command))
        {
            return Unavailable(
                MacCapabilityTier.Unsupported,
                command,
                "a native Apple Events range candidate that has not completed real CLI and MCP acceptance");
        }

        var separator = command.IndexOf('.');
        var category = separator > 0 ? command[..separator] : command;
        var action = separator > 0 ? command[(separator + 1)..] : string.Empty;

        if (MacOfficeActionCatalog.TryGet(command, out _))
        {
            if (officeCandidateEnabled)
            {
                return new MacCommandCapability(
                    true,
                    MacCapabilityTier.OfficeAddIn,
                    string.Empty);
            }
            return Unavailable(
                MacCapabilityTier.OfficeAddIn,
                command,
                "an Office.js implementation candidate that has not passed live contract parity validation");
        }

        if (category == "powerquery")
        {
            if (action is "list"
                or "view"
                or "get-load-config"
                or "update")
            {
                return new MacCommandCapability(
                    true,
                    MacCapabilityTier.PowerQueryPackage,
                    string.Empty);
            }

            return new MacCommandCapability(
                true,
                MacCapabilityTier.MacroHelper,
                string.Empty);
        }

        if (category == "vba")
        {
            return action == "run"
                ? Unavailable(
                    MacCapabilityTier.MacroHelper,
                    command,
                    "the optional macro helper tier, which requires the user to enable macros")
                : Unavailable(
                    MacCapabilityTier.VbaProjectModel,
                    command,
                    "the optional macOS VBA project-model tier; preflight reports that " +
                    $"{MacVbaPreflight.DescribeProjectModel(vbaPreflight.ProjectModelAccess)}, " +
                    "and Excel's installed scripting dictionary exposes no VBA project-model route");
        }

        if (command is "analysis.create-scenario" or "analysis.show-scenario")
        {
            return Unavailable(
                MacCapabilityTier.MacroHelper,
                command,
                "the optional trusted VBA helper because the installed native dictionary does not expose " +
                "exact scenario creation or Scenario.Show");
        }

        if (category is "connection" or "querytable" or "analysis" or "pythoninexcel")
        {
            return Unavailable(
                MacCapabilityTier.Native,
                command,
                "an Apple Events route whose exact result, completion, error, and cleanup semantics " +
                "have not yet passed a prompt-free real-Excel fixture");
        }

        if (category == "screenshot")
        {
            return Unavailable(
                MacCapabilityTier.OptionalNativeHelper,
                command,
                "an optional native screen-capture helper with explicit Screen Recording permission");
        }

        if (category is "table"
            or "tablecolumn"
            or "chart"
            or "chartconfig"
            or "pivottable"
            or "pivottablefield"
            or "pivottablecalc"
            or "conditionalformat"
            or "drawing"
            or "slicer"
            or "rangeformat")
        {
            return Unavailable(
                MacCapabilityTier.OfficeAddIn,
                command,
                "an Office.js implementation that has not passed contract parity validation");
        }

        return Unavailable(
            MacCapabilityTier.Unsupported,
            command,
            false,
            MacCapabilityTier.Unsupported,
            $"Command '{command}' requires a capability that is not supported by the macOS Excel backend. It remains available on Windows.",
            "No generated Windows contract exists for this command.",
            [],
            "NotTested",
            "The command is absent from the generated public action inventory.",
            "Unverified.",
            "a capability that is not supported by the macOS Excel backend");
    }

    private static MacCommandCapability ToCapability(MacActionInventoryItem item)
    {
        if (!Enum.TryParse<MacCapabilityTier>(item.Tier, out var tier))
        {
            throw new InvalidOperationException(
                $"Generated macOS capability tier '{item.Tier}' for '{item.Command}' is invalid.");
        }

        return new MacCommandCapability(
            item.Command,
            item.IsAvailable,
            tier,
            item.IsAvailable
                ? string.Empty
                : $"Command '{item.Command}' is unavailable on macOS: " +
                  $"{item.Blocker.TrimEnd('.')}. It remains available on Windows.",
            item.WindowsSemantics,
            item.WindowsVariants,
            item.ImplementationStatus,
            item.Evidence,
            item.ExcelApiVersion,
            item.Blocker);
    }
}
