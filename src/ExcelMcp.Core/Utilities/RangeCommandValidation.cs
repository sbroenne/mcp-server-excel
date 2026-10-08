using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Commands.Range;

namespace Sbroenne.ExcelMcp.Core.Utilities;

internal static class RangeCommandValidation
{
    internal const int ConflictExampleLimit = 10;

    internal static void ValidateOverwritePolicy(OverwritePolicy overwritePolicy)
    {
        if (!Enum.IsDefined(overwritePolicy))
            throw new ArgumentOutOfRangeException(nameof(overwritePolicy), overwritePolicy, "Unknown overwrite policy.");
    }

    internal static void ValidateFormulaReferenceStyle(FormulaReferenceStyle referenceStyle)
    {
        if (!Enum.IsDefined(referenceStyle))
            throw new ArgumentOutOfRangeException(nameof(referenceStyle));
    }

    internal static void ValidateDimensions<T>(List<List<T>> payload, int rows, int columns, string parameterName, string itemType)
    {
        ValidateRowWidths(payload, columns, parameterName, itemType);
        if (payload.Count != rows)
            throw new ArgumentException(
                $"{itemType} array row count ({payload.Count}) doesn't match range row count ({rows}).", parameterName);
    }

    internal static void ValidateRowWidths<T>(List<List<T>> payload, int columns, string parameterName, string itemType)
    {
        for (var row = 0; row < payload.Count; row++)
            if (payload[row].Count != columns)
                throw new ArgumentException(
                    $"{itemType} array row {row + 1} column count ({payload[row].Count}) doesn't match range column count ({columns})",
                    parameterName);
    }

    internal static bool IsOccupied(object? value, object? formula) =>
        value is not null || formula is string text && text.StartsWith('=');

    internal static void ThrowOccupiedDestination(string sheetName, List<string> conflicts, bool truncated) =>
        throw new OperationFailureException(
            OperationFailureCategory.Conflict,
            $"Cannot write to occupied cells on sheet '{sheetName}' with overwrite_policy='reject-nonempty'. " +
            $"Conflicting addresses: {string.Join(", ", conflicts)}" +
            (truncated ? " (additional conflicts not listed)." : ".") +
            " No write was attempted. Choose an empty destination, or use overwrite_policy='allow' " +
            "(CLI: --overwrite-policy allow) only when replacing existing content is authorized.");

    internal static void ThrowMergedCellWriteError(string rangeAddress, List<string> mergedRanges) =>
        throw new OperationFailureException(
            OperationFailureCategory.Conflict,
            $"Cannot write to range '{rangeAddress}' because the write intersects merged cells. " +
            $"{(mergedRanges.Count == 1 ? "Merged range" : "Merged ranges")}: {string.Join(", ", mergedRanges)}. " +
            "Write only to each merged range's top-left cell, or unmerge the affected range before writing.");

    internal static string ColumnLetter(int column)
    {
        var name = string.Empty;
        while (column > 0)
        {
            column--;
            name = Convert.ToChar('A' + column % 26) + name;
            column /= 26;
        }
        return name;
    }
}
