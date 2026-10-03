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
    /// <summary>
    /// Visits every unique cell in an already-resolved native range.
    /// The callback borrows the cell; this method owns and releases its reference.
    /// </summary>
    internal static void VisitCells(Excel.Range range, CancellationToken ct, Action<Excel.Range> visit)
    {
        Excel.Areas? areas = null;
        try
        {
            areas = range.Areas;
            for (int index = 1; index <= areas.Count; index++)
            {
                Excel.Range? area = null;
                Excel.Range? cells = null;
                Excel.Range? rows = null;
                Excel.Range? columns = null;
                try
                {
                    area = areas[index];
                    cells = area.Cells;
                    rows = area.Rows;
                    columns = area.Columns;
                    int rowCount = rows.Count;
                    int columnCount = columns.Count;
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int column = 1; column <= columnCount; column++)
                        {
                            ct.ThrowIfCancellationRequested();
                            Excel.Range? cell = null;
                            try
                            {
                                cell = cells[row, column];
                                visit(cell);
                            }
                            finally
                            {
                                ComUtilities.Release(ref cell);
                            }
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref columns);
                    ComUtilities.Release(ref rows);
                    ComUtilities.Release(ref cells);
                    ComUtilities.Release(ref area);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref areas);
        }
    }

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
            var areaAddresses = ParseSupportedRangeAreas(rangeAddress);
            if (areaAddresses is null)
            {
                specificError = $"Sheet '{sheetName}' exists, but range '{rangeAddress}' is invalid. " +
                               $"Verify the range address format (e.g., 'A1:E10', 'A1', 'A:A').";
                throw new OperationFailureException(
                    OperationFailureCategory.InvalidInput,
                    specificError);
            }

            return ResolveRangeAreas(sheet, areaAddresses);
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
        int workbookBangIndex = workbookName.LastIndexOf('!');
        int requestedBangIndex = requestedName.LastIndexOf('!');
        if (workbookBangIndex < 0 || requestedBangIndex < 0)
        {
            return workbookBangIndex == requestedBangIndex
                && string.Equals(
                    workbookName,
                    requestedName,
                    StringComparison.OrdinalIgnoreCase);
        }

        return string.Equals(
                workbookName[..workbookBangIndex].Trim('\''),
                requestedName[..requestedBangIndex].Trim('\''),
                StringComparison.OrdinalIgnoreCase)
            && string.Equals(
                workbookName[(workbookBangIndex + 1)..],
                requestedName[(requestedBangIndex + 1)..],
                StringComparison.OrdinalIgnoreCase);
    }

    private static Excel.Range ResolveRangeAreas(
        Excel.Worksheet sheet, List<string> areaAddresses)
    {
        Excel.Application? app = null;
        Excel.Range? result = null;
        try
        {
            if (areaAddresses.Count > 1)
            {
                app = sheet.Application;
            }
            foreach (var address in areaAddresses)
            {
                Excel.Range? area = null;
                Excel.Range? combined = null;
                try
                {
                    // Native unions avoid Excel's locale-dependent address separator.
                    area = sheet.Range[address];
                    if (result is null)
                    {
                        result = area;
                        area = null;
                    }
                    else
                    {
                        combined = app!.Union(result, area);
                        ComUtilities.Release(ref result);
                        result = combined;
                        combined = null;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref combined);
                    ComUtilities.Release(ref area);
                }
            }
            var resolved = result ?? throw new InvalidOperationException("No range areas were resolved.");
            result = null;
            return resolved;
        }
        finally
        {
            ComUtilities.Release(ref result);
            ComUtilities.Release(ref app);
        }
    }

    private static List<string>? ParseSupportedRangeAreas(string rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(rangeAddress))
        {
            return null;
        }

        List<string> areas = [];
        int areaStart = 0;
        int bracketDepth = 0;
        for (int index = 0; index < rangeAddress.Length; index++)
        {
            char character = rangeAddress[index];
            if (character == '\'' && bracketDepth > 0 && index + 1 < rangeAddress.Length)
            {
                index++;
                continue;
            }

            if (character == '[')
            {
                bracketDepth++;
            }
            else if (character == ']')
            {
                if (--bracketDepth < 0)
                {
                    return null;
                }
            }
            else if (character == ',' && bracketDepth == 0)
            {
                var area = rangeAddress[areaStart..index].Trim();
                if (!IsSupportedRangeArea(area))
                {
                    return null;
                }
                areas.Add(area);
                areaStart = index + 1;
            }
        }

        var lastArea = rangeAddress[areaStart..].Trim();
        if (bracketDepth != 0 || !IsSupportedRangeArea(lastArea))
        {
            return null;
        }
        areas.Add(lastArea);
        return areas;
    }

    private static bool IsSupportedRangeArea(string area)
    {
        if (TryParseCellReference(area, out _, out _))
        {
            return true;
        }

        if (area.EndsWith('#'))
        {
            return TryParseCellReference(area[..^1], out _, out _);
        }

        if (IsStructuredReference(area))
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

    private static bool IsStructuredReference(string area)
    {
        int firstBracket = area.IndexOf('[');
        if (firstBracket <= 0
            || area[^1] != ']'
            || area[..firstBracket].Any(character =>
                char.IsWhiteSpace(character) || character is '[' or ']' or '#'))
        {
            return false;
        }

        var hasContent = new List<bool>();
        for (int index = firstBracket; index < area.Length; index++)
        {
            char character = area[index];
            if (character == '[')
            {
                hasContent.Add(false);
            }
            else if (character == ']')
            {
                if (hasContent.Count == 0 || !hasContent[^1])
                {
                    return false;
                }

                hasContent.RemoveAt(hasContent.Count - 1);
                if (hasContent.Count == 0)
                {
                    return index == area.Length - 1;
                }

                hasContent[^1] = true;
            }
            else
            {
                if (hasContent.Count == 0)
                {
                    return false;
                }

                if (character == '\'' && index + 1 < area.Length)
                {
                    index++;
                    hasContent[^1] = true;
                }
                else if (!char.IsWhiteSpace(character) && character != ',')
                {
                    hasContent[^1] = true;
                }
            }
        }

        return false;
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
    private static RangeCopyResult CopyRange(
        IExcelBatch batch,
        string sourceSheet,
        string sourceRange,
        string targetSheet,
        string targetRange,
        PasteKind pasteKind,
        bool transpose,
        bool skipBlanks,
        OverwritePolicy overwritePolicy)
    {
        ValidateOverwritePolicy(overwritePolicy);
        var pasteType = pasteKind switch
        {
            PasteKind.All => Excel.XlPasteType.xlPasteAll,
            PasteKind.Values => Excel.XlPasteType.xlPasteValues,
            PasteKind.Formulas => Excel.XlPasteType.xlPasteFormulas,
            PasteKind.Formats => Excel.XlPasteType.xlPasteFormats,
            PasteKind.Validation => Excel.XlPasteType.xlPasteValidation,
            _ => throw new ArgumentOutOfRangeException(nameof(pasteKind))
        };
        var result = new RangeCopyResult
        {
            FilePath = batch.WorkbookPath,
            Action = "copy",
            SourceSheet = sourceSheet,
            TargetSheet = targetSheet,
            PasteKind = pasteKind,
            Transpose = transpose,
            SkipBlanks = skipBlanks
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? srcRange = null;
            Excel.Range? tgtRange = null;
            Excel.Range? destination = null;
            Excel.Range? sourceCells = null;
            bool copyStarted = false;
            Exception? primaryFailure = null;
            try
            {
                srcRange = RangeHelpers.ResolveRange(ctx.Book, sourceSheet, sourceRange);
                tgtRange = RangeHelpers.ResolveRange(ctx.Book, targetSheet, targetRange);
                destination = ResolveCopyDestination(srcRange!, tgtRange!, transpose);
                result.SourceAddress = srcRange!.Address;
                result.DestinationAddress = destination.Address;
                if (pasteKind is PasteKind.All or PasteKind.Values or PasteKind.Formulas
                    && overwritePolicy == OverwritePolicy.RejectNonempty)
                {
                    if (skipBlanks)
                    {
                        sourceCells = srcRange.Cells;
                        var sourceSize = GetContentDimensions(srcRange);
                        Dictionary<(int Row, int Column), bool> content = [];
                        EnsureDestinationWritable(ctx, destination, overwritePolicy, ct, (row, column) =>
                        {
                            var key = transpose
                                ? (column % sourceSize.Rows, row % sourceSize.Columns)
                                : (row % sourceSize.Rows, column % sourceSize.Columns);
                            if (!content.TryGetValue(key, out bool writes))
                            {
                                Excel.Range? sourceCell = null;
                                try
                                {
                                    sourceCell = sourceCells[key.Item1 + 1, key.Item2 + 1];
                                    object? value = sourceCell.Value2;
                                    object hasFormula = sourceCell.HasFormula;
                                    if (hasFormula is not bool formula)
                                    {
                                        throw new InvalidOperationException(
                                            "Cannot inspect source blanks: Excel returned an indeterminate formula state. No write was attempted.");
                                    }
                                    writes = value is not null || formula;
                                    content.Add(key, writes);
                                }
                                finally
                                {
                                    ComUtilities.Release(ref sourceCell);
                                }
                            }
                            return writes;
                        });
                    }
                    else
                    {
                        EnsureDestinationWritable(ctx, destination, overwritePolicy, ct);
                    }
                }

                ct.ThrowIfCancellationRequested();
                copyStarted = true;
                srcRange.Copy();
                destination.PasteSpecial(pasteType, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                    skipBlanks, transpose);
            }
            catch (Exception ex)
            {
                primaryFailure = ex;
            }
            finally
            {
                try
                {
                    if (copyStarted)
                    {
                        try
                        {
                            ctx.App.CutCopyMode = (Excel.XlCutCopyMode)0;
                        }
                        catch (Exception cleanupFailure)
                        {
                            primaryFailure = primaryFailure is null
                                ? cleanupFailure
                                : new AggregateException("Copy failed and Excel's copy mode could not be cleared.",
                                    primaryFailure, cleanupFailure);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref sourceCells);
                    ComUtilities.Release(ref destination);
                    ComUtilities.Release(ref tgtRange);
                    ComUtilities.Release(ref srcRange);
                }
            }
            if (primaryFailure is not null)
            {
                ExceptionDispatchInfo.Capture(primaryFailure).Throw();
            }
            result.Success = true;
            return result;
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
