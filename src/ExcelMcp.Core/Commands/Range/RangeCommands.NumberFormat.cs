using System.Globalization;
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

                // Set uniform number format for entire range
                // Use NumberFormatLocal to interpret format codes according to user's locale
                range.NumberFormatLocal = formatCode;

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
    /// Sets number format using structured options (locale-aware).
    /// This method automatically generates the correct locale-specific format code.
    /// </summary>
    /// <param name="batch">The Excel batch operation context</param>
    /// <param name="sheetName">Sheet name (empty for named ranges)</param>
    /// <param name="rangeAddress">Range address or named range name</param>
    /// <param name="options">Structured format options</param>
    /// <returns>Operation result</returns>
    public OperationResult SetNumberFormatStructured(IExcelBatch batch, string sheetName, string rangeAddress, NumberFormatOptions options)
    {
        // Build locale-aware format code from structured options using current culture
        var formatCode = LocaleAwareFormatBuilder.BuildFormatCode(options, CultureInfo.CurrentCulture);

        // Use the existing method with the generated format code
        return SetNumberFormat(batch, sheetName, rangeAddress, formatCode);
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
                                cell.NumberFormatLocal = formats[row - 1][col - 1];
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
                            formatArray[row, col] = formats[row][col];
                        }
                    }

                    // Set number formats via 2D array
                    range.NumberFormatLocal = formatArray;
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

