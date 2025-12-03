using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// RangeCommands partial class - Number formatting operations
/// </summary>
public partial class RangeCommands
{
    // === NUMBER FORMAT OPERATIONS ===

    /// <inheritdoc />
    public RangeNumberFormatResult GetNumberFormats(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new RangeNumberFormatResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                // Get actual address from Excel
                result.RangeAddress = range.Address;

                // Get number formats - Excel COM behavior:
                // - Single cell: returns string
                // - Multiple cells, all same format: returns string
                // - Multiple cells, mixed formats: returns DBNull (must read cell-by-cell)
                object numberFormats = range.NumberFormat;

                // Get dimensions
                int rowCount = Convert.ToInt32(range.Rows.Count);
                int columnCount = Convert.ToInt32(range.Columns.Count);

                result.RowCount = rowCount;
                result.ColumnCount = columnCount;

                // Check if we have mixed formats (DBNull or null)
                if (numberFormats is null or DBNull)
                {
                    // Mixed formats - must read cell-by-cell
                    dynamic? cells = null;
                    try
                    {
                        cells = range.Cells;
                        for (int row = 1; row <= rowCount; row++)
                        {
                            var rowList = new List<string>();
                            for (int col = 1; col <= columnCount; col++)
                            {
                                dynamic? cell = null;
                                try
                                {
                                    cell = cells[row, col];
                                    var format = cell.NumberFormat?.ToString() ?? "General";
                                    rowList.Add(format);
                                }
                                finally
                                {
                                    ComUtilities.Release(ref cell);
                                }
                            }
                            result.Formats.Add(rowList);
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref cells);
                    }
                }
                else if (numberFormats is string formatStr)
                {
                    // All cells have same format
                    for (int row = 0; row < rowCount; row++)
                    {
                        var rowList = new List<string>();
                        for (int col = 0; col < columnCount; col++)
                        {
                            rowList.Add(formatStr);
                        }
                        result.Formats.Add(rowList);
                    }
                }
                else
                {
                    // Should be a 2D array (rare case)
                    object[,] formats = (object[,])numberFormats;
                    for (int row = 0; row < rowCount; row++)
                    {
                        var rowList = new List<string>();
                        for (int col = 0; col < columnCount; col++)
                        {
                            var format = formats[row, col]?.ToString() ?? "General";
                            rowList.Add(format);
                        }
                        result.Formats.Add(rowList);
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetNumberFormat(IExcelBatch batch, string sheetName, string rangeAddress, string formatCode)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "set-number-format"
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                // Automatically add US LCID prefix for date/time formats to ensure cross-culture compatibility.
                // This prevents 'm' being interpreted as minutes instead of months on non-US locales.
                string effectiveFormat = EnsureDateTimeLcid(formatCode);
                range.NumberFormat = effectiveFormat;

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
    /// Ensures date/time format codes have a US LCID prefix for cross-culture compatibility.
    /// If the format contains date/time specifiers (d, m, y, h, s) but no LCID prefix,
    /// automatically adds [$-409] to ensure consistent interpretation across all locales.
    /// </summary>
    /// <param name="formatCode">The original format code.</param>
    /// <returns>Format code with LCID prefix if needed, or original if not a date/time format.</returns>
    private static string EnsureDateTimeLcid(string formatCode)
    {
        if (string.IsNullOrEmpty(formatCode))
            return formatCode;

        // Already has LCID prefix - don't modify
        if (formatCode.Contains("[$-"))
            return formatCode;

        // Check if this looks like a date/time format (contains d, m, y, h, s outside of quotes)
        // but NOT a number format (which uses # , 0 . %)
        if (!ContainsDateTimeSpecifiers(formatCode))
            return formatCode;

        // Add US LCID prefix for cross-culture compatibility
        return $"[$-409]{formatCode}";
    }

    /// <summary>
    /// Checks if a format code contains date/time specifiers (d, m, y, h, s) outside of quoted strings.
    /// Returns false for pure number formats (containing only #, 0, ., ,, %, $, etc.)
    /// </summary>
    private static bool ContainsDateTimeSpecifiers(string formatCode)
    {
        bool inQuotes = false;
        bool hasDateTimeChars = false;
        bool hasNumberChars = false;

        foreach (char c in formatCode)
        {
            if (c == '"')
            {
                inQuotes = !inQuotes;
                continue;
            }

            if (inQuotes)
                continue;

            char lower = char.ToLowerInvariant(c);

            // Date/time specifiers
            if (lower is 'd' or 'm' or 'y' or 'h' or 's')
            {
                hasDateTimeChars = true;
            }
            // Number format specifiers (not date/time)
            else if (c is '#' or '0' or '%')
            {
                hasNumberChars = true;
            }
        }

        // Only treat as date/time if it has date/time chars and is NOT primarily a number format
        // Exception: "m" alone with number chars (like #,##0) is minutes in time context but we skip those
        return hasDateTimeChars && !hasNumberChars;
    }

    /// <inheritdoc />
    public OperationResult SetNumberFormats(IExcelBatch batch, string sheetName, string rangeAddress, List<List<string>> formats)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "set-number-formats"
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                int rowCount = Convert.ToInt32(range.Rows.Count);
                int columnCount = Convert.ToInt32(range.Columns.Count);

                // Validate dimensions match
                if (formats.Count != rowCount)
                {
                    throw new ArgumentException($"Format array row count ({formats.Count}) doesn't match range row count ({rowCount})", nameof(formats));
                }

                for (int i = 0; i < formats.Count; i++)
                {
                    if (formats[i].Count != columnCount)
                    {
                        throw new ArgumentException($"Format array row {i + 1} column count ({formats[i].Count}) doesn't match range column count ({columnCount})", nameof(formats));
                    }
                }

                // If single row or column, can't use 2D array - must set cell by cell
                if (rowCount == 1 || columnCount == 1)
                {
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= columnCount; col++)
                        {
                            dynamic? cell = null;
                            try
                            {
                                cell = range.Cells[row, col];
                                // Auto-add LCID for date/time formats
                                cell.NumberFormat = EnsureDateTimeLcid(formats[row - 1][col - 1]);
                            }
                            finally
                            {
                                ComUtilities.Release(ref cell);
                            }
                        }
                    }
                }
                else
                {
                    // For multi-row, multi-column ranges, Excel COM expects 1-based 2D array
                    object[,] formatArray = new object[rowCount, columnCount];
                    for (int row = 0; row < rowCount; row++)
                    {
                        for (int col = 0; col < columnCount; col++)
                        {
                            // Auto-add LCID for date/time formats
                            formatArray[row, col] = EnsureDateTimeLcid(formats[row][col]);
                        }
                    }

                    // Set number formats via 2D array
                    range.NumberFormat = formatArray;
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }
}

