namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacRangeAcceptanceCatalog
{
    private static readonly HashSet<string> Candidates = new(StringComparer.Ordinal)
    {
        "range.copy",
        "range.copy-values",
        "range.copy-formulas",
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

    public static bool IsAcceptanceEnabled(string command, string? environmentValue) =>
        environmentValue == "1" && Candidates.Contains(command);

    public static bool Contains(string command) => Candidates.Contains(command);
}
