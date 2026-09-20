using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;


namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range resolution and helper methods for RangeCommands
/// </summary>
public static class RangeHelpers
{
    private const int ExcelMaxRows = 1_048_576;
    private const int ExcelMaxColumns = 16_384;

    private static readonly Regex CellReferenceRegex = new(
        @"^\$?(?<column>[A-Z]{1,3})\$?(?<row>\d{1,7})$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex ColumnReferenceRegex = new(
        @"^\$?(?<column>[A-Z]{1,3})$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex RowReferenceRegex = new(
        @"^\$?(?<row>\d{1,7})$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <summary>
    /// Resolves a range address to a Range COM object.
    /// Supports both regular ranges (Sheet1!A1:D10) and named ranges.
    /// Throws a categorized failure when the sheet, named range, or address cannot be resolved.
    /// </summary>
    public static dynamic? ResolveRange(dynamic book, string sheetName, string rangeAddress, out string? specificError)
    {
        specificError = null;

        // Named range (empty sheetName)
        if (string.IsNullOrEmpty(sheetName))
        {
            dynamic? names = null;
            dynamic? name = null;
            dynamic? refersToRange = null;
            try
            {
                names = book.Names;
                name = ResolveNameByExactKey(names, rangeAddress);
                refersToRange = name.RefersToRange;
                var result = refersToRange;
                refersToRange = null;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref refersToRange);
                ComUtilities.Release(ref name);
                ComUtilities.Release(ref names);
            }
        }

        // Regular range (sheet + address)
        // First check if sheet exists
        Excel.Worksheet? sheet = null;
        try
        {
            sheet = ComUtilities.FindSheet(book, sheetName);
            if (sheet == null)
            {
                specificError = $"Sheet '{sheetName}' not found.";
                throw new OperationFailureException(
                    OperationFailureCategory.NotFound,
                    specificError);
            }

            // Sheet exists, now try to get the range
            if (!IsSupportedRangeAddress(rangeAddress))
            {
                specificError = $"Sheet '{sheetName}' exists, but range '{rangeAddress}' is invalid. " +
                               $"Verify the range address format (e.g., 'A1:E10', 'A1', 'A:A').";
                throw new OperationFailureException(
                    OperationFailureCategory.InvalidInput,
                    specificError);
            }

            return sheet.Range[rangeAddress];
        }
        finally
        {
            ComUtilities.Release(ref sheet);
        }
    }

    /// <summary>
    /// Resolves a range address to a Range COM object (backward compatibility).
    /// Supports both regular ranges (Sheet1!A1:D10) and named ranges.
    /// </summary>
    public static dynamic? ResolveRange(dynamic book, string sheetName, string rangeAddress)
    {
        string? ignoredError;
        return ResolveRange(book, sheetName, rangeAddress, out ignoredError);
    }

    /// <summary>
    /// Gets appropriate error message for range resolution failure
    /// </summary>
    public static string GetResolveError(string sheetName, string rangeAddress)
    {
        if (string.IsNullOrEmpty(sheetName))
        {
            return $"Named range '{rangeAddress}' not found.";
        }
        return $"Sheet '{sheetName}' or range '{rangeAddress}' not found.";
    }

    private static dynamic ResolveNameByExactKey(dynamic names, string rangeAddress)
    {
        dynamic? name = null;
        try
        {
            name = names.Item(rangeAddress);
            var result = name;
            name = null;
            return result;
        }
        catch (COMException ex)
        {
            if (WorkbookContainsName(names, rangeAddress))
            {
                ExceptionDispatchInfo.Capture(ex).Throw();
            }

            throw new OperationFailureException(
                OperationFailureCategory.NotFound,
                $"Named range '{rangeAddress}' not found.",
                ex);
        }
        finally
        {
            ComUtilities.Release(ref name);
        }
    }

    private static bool WorkbookContainsName(dynamic names, string rangeAddress)
    {
        int count = Convert.ToInt32(names.Count);
        for (int index = 1; index <= count; index++)
        {
            dynamic? currentName = null;
            try
            {
                currentName = names.Item(index);
                string currentNameText = currentName.Name?.ToString() ?? string.Empty;
                if (NameMatches(currentNameText, rangeAddress))
                {
                    return true;
                }
            }
            finally
            {
                ComUtilities.Release(ref currentName);
            }
        }

        return false;
    }

    private static bool NameMatches(string workbookName, string requestedName)
    {
        if (string.Equals(workbookName, requestedName, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        int bangIndex = workbookName.LastIndexOf('!');
        return bangIndex >= 0
            && string.Equals(
                workbookName[(bangIndex + 1)..].Trim('\''),
                requestedName,
                StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSupportedRangeAddress(string rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(rangeAddress))
        {
            return false;
        }

        foreach (var area in rangeAddress.Split(','))
        {
            string trimmedArea = area.Trim();
            if (!IsSupportedRangeArea(trimmedArea)
                && !IsExtendedRangeReference(trimmedArea))
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsExtendedRangeReference(string area) =>
        area.EndsWith('#')
        || (area.Contains('[') && area.Contains(']'));

    private static bool IsSupportedRangeArea(string area)
    {
        if (TryParseCellReference(area, out _, out _))
        {
            return true;
        }

        var parts = area.Split(':');
        if (parts.Length != 2)
        {
            return false;
        }

        return AreSupportedCellBounds(parts[0], parts[1])
            || AreSupportedColumnBounds(parts[0], parts[1])
            || AreSupportedRowBounds(parts[0], parts[1]);
    }

    private static bool AreSupportedCellBounds(string start, string end) =>
        TryParseCellReference(start, out var startColumn, out var startRow)
        && TryParseCellReference(end, out var endColumn, out var endRow)
        && startColumn >= 1
        && endColumn >= 1
        && startRow >= 1
        && endRow >= 1;

    private static bool AreSupportedColumnBounds(string start, string end) =>
        TryParseColumnReference(start, out var startColumn)
        && TryParseColumnReference(end, out var endColumn)
        && startColumn >= 1
        && endColumn >= 1;

    private static bool AreSupportedRowBounds(string start, string end) =>
        TryParseRowReference(start, out var startRow)
        && TryParseRowReference(end, out var endRow)
        && startRow >= 1
        && endRow >= 1;

    private static bool TryParseCellReference(string reference, out int column, out int row)
    {
        column = 0;
        row = 0;
        var match = CellReferenceRegex.Match(reference);
        if (!match.Success)
        {
            return false;
        }

        return TryParseColumn(match.Groups["column"].Value, out column)
            && int.TryParse(match.Groups["row"].Value.TrimStart('$'), out row)
            && row is >= 1 and <= ExcelMaxRows;
    }

    private static bool TryParseColumnReference(string reference, out int column)
    {
        column = 0;
        var match = ColumnReferenceRegex.Match(reference);
        return match.Success && TryParseColumn(match.Groups["column"].Value, out column);
    }

    private static bool TryParseRowReference(string reference, out int row)
    {
        row = 0;
        var match = RowReferenceRegex.Match(reference);
        return match.Success
            && int.TryParse(match.Groups["row"].Value.TrimStart('$'), out row)
            && row is >= 1 and <= ExcelMaxRows;
    }

    private static bool TryParseColumn(string columnText, out int column)
    {
        column = 0;
        foreach (char character in columnText.TrimStart('$').ToUpperInvariant())
        {
            if (character is < 'A' or > 'Z')
            {
                return false;
            }

            column = (column * 26) + character - 'A' + 1;
        }

        return column is >= 1 and <= ExcelMaxColumns;
    }

    /// <summary>
    /// Converts a value to a proper Excel cell value, handling System.Text.Json.JsonElement.
    /// MCP framework deserializes JSON arrays to JsonElement objects which cannot be marshalled to COM Variant.
    /// This helper detects JsonElement and converts to proper C# types before COM assignment.
    /// </summary>
    /// <param name="value">Value from MCP JSON deserialization or direct C# types</param>
    /// <returns>Proper C# type (string, long, double, bool) for COM marshalling</returns>
    public static object ConvertToCellValue(object? value)
    {
        if (value == null)
            return string.Empty;

        // Handle System.Text.Json.JsonElement (from MCP JSON deserialization)
        if (value is System.Text.Json.JsonElement jsonElement)
        {
            return jsonElement.ValueKind switch
            {
                System.Text.Json.JsonValueKind.String => jsonElement.GetString() ?? string.Empty,
                System.Text.Json.JsonValueKind.Number => jsonElement.TryGetInt64(out var i64) ? i64 : jsonElement.GetDouble(),
                System.Text.Json.JsonValueKind.True => true,
                System.Text.Json.JsonValueKind.False => false,
                System.Text.Json.JsonValueKind.Null => string.Empty,
                _ => jsonElement.ToString() ?? string.Empty
            };
        }

        // Already a proper type (from CLI or tests)
        return value;
    }
}

/// <summary>
/// Internal helper methods for RangeCommands partial class
/// </summary>
public partial class RangeCommands
{
    /// <summary>
    /// Helper for clear operations (resolve range, apply action, release)
    /// </summary>
    private static OperationResult ClearRange(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        string action,
        Action<dynamic> clearAction)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = action };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress);
                if (range == null)
                {
                    throw new InvalidOperationException(RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                clearAction(range);
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <summary>
    /// Helper for copy operations (resolve source + target ranges, apply copy action, release both)
    /// </summary>
    private static OperationResult CopyRange(
        IExcelBatch batch,
        string sourceSheet,
        string sourceRange,
        string targetSheet,
        string targetRange,
        string action,
        Action<dynamic, dynamic> copyAction)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = action };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? srcRange = null;
            dynamic? tgtRange = null;
            try
            {
                srcRange = RangeHelpers.ResolveRange(ctx.Book, sourceSheet, sourceRange, out string? srcError);
                if (srcRange == null)
                {
                    throw new InvalidOperationException(srcError ?? RangeHelpers.GetResolveError(sourceSheet, sourceRange));
                }

                tgtRange = RangeHelpers.ResolveRange(ctx.Book, targetSheet, targetRange, out string? tgtError);
                if (tgtRange == null)
                {
                    throw new InvalidOperationException(tgtError ?? RangeHelpers.GetResolveError(targetSheet, targetRange));
                }

                copyAction(srcRange, tgtRange);
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref srcRange);
                ComUtilities.Release(ref tgtRange);
            }
        });
    }

    /// <summary>
    /// Helper for insert/delete row/column operations (resolve range, get EntireRow/EntireColumn, apply action, release)
    /// </summary>
    private static OperationResult ModifyRowsOrColumns(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        string action,
        Func<dynamic, dynamic> accessor,
        Action<dynamic> operation)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = action };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            dynamic? target = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                target = accessor(range);
                operation(target);
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref range);
            }
        });
    }
}
