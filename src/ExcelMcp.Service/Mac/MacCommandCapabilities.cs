namespace Sbroenne.ExcelMcp.Service.Mac;

internal enum MacCapabilityTier
{
    Native,
    OfficeAddIn,
    PowerQueryPackage,
    MacroHelper,
    VbaProjectModel,
    OptionalNativeHelper,
    Unsupported
}

internal sealed record MacCommandCapability(
    bool IsAvailable,
    MacCapabilityTier RequiredTier,
    string UnavailableMessage);

internal static class MacCommandCapabilities
{
    private static readonly HashSet<string> NativeCommands = new(StringComparer.Ordinal)
    {
        "sheet.list",
        "sheet.create",
        "sheet.rename",
        "sheet.delete",
        "sheet.set-visibility",
        "sheet.get-visibility",
        "sheet.show",
        "sheet.hide",
        "sheet.very-hide",
        "sheet.set-tab-color",
        "sheet.get-tab-color",
        "sheet.clear-tab-color",
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
        "calculation.calculate",
        "analysis.goal-seek",
        "analysis.create-data-table"
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
        "conditionalformat",
        "drawing",
        "slicer"
    };

    public static MacCommandCapability Get(
        string command,
        MacVbaPreflightResult? vbaPreflight = null)
    {
        if (NativeCommands.Contains(command))
        {
            return new MacCommandCapability(true, MacCapabilityTier.Native, string.Empty);
        }

        var separator = command.IndexOf('.');
        var category = separator > 0 ? command[..separator] : command;
        var action = separator > 0 ? command[(separator + 1)..] : string.Empty;

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

            return Unavailable(
                MacCapabilityTier.PowerQueryPackage,
                command,
                "the secure saved-package Power Query mutation tier, which is not enabled in this release");
        }

        if (category == "vba")
        {
            vbaPreflight ??= MacVbaPreflight.Check();
            return action == "run"
                ? Unavailable(
                    MacCapabilityTier.MacroHelper,
                    command,
                    "the optional macOS macro execution tier; preflight reports that " +
                    $"{MacVbaPreflight.DescribeMacroExecution(vbaPreflight.MacroExecution)}; " +
                    "a repository-owned synthetic fixture has not yet proven unattended, " +
                    "workbook-qualified execution through both CLI and MCP")
                : Unavailable(
                    MacCapabilityTier.VbaProjectModel,
                    command,
                    "the optional macOS VBA project-model tier; preflight reports that " +
                    $"{MacVbaPreflight.DescribeProjectModel(vbaPreflight.ProjectModelAccess)}, " +
                    "and Excel's installed scripting dictionary exposes no VBA project-model route");
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

        if (OfficeAddInCategories.Contains(category)
            || category == "rangeformat")
        {
            return Unavailable(
                MacCapabilityTier.OfficeAddIn,
                command,
                "the optional Office.js add-in tier, which is not installed in this release");
        }

        return Unavailable(
            MacCapabilityTier.Unsupported,
            command,
            "a capability that is not supported by the macOS Excel backend");
    }

    private static MacCommandCapability Unavailable(
        MacCapabilityTier tier,
        string command,
        string requirement) =>
        new(
            false,
            tier,
            $"Command '{command}' requires {requirement}. It remains available on Windows.");
}
