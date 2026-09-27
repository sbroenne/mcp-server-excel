using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Typed NumberFormat reads return invariant codes. Explicit NumberFormatLocal writes preserve
/// currency literals while translating the invariant codes accepted by both entry points.
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
                object numberFormats = ((Excel.Range)range).NumberFormat;

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
                                    var format = ((Excel.Range)cell).NumberFormat?.ToString() ?? "General";
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
                    // Should be a 2D array (rare case). COM arrays are 1-based.
                    object[,] formats = (object[,])numberFormats;
                    for (int row = 1; row <= rowCount; row++)
                    {
                        var rowList = new List<string>();
                        for (int col = 1; col <= columnCount; col++)
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

                ((Excel.Range)range).NumberFormatLocal = ctx.FormatTranslator.TranslateToLocale(formatCode);

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
    public OperationResult SetNumberFormats(IExcelBatch batch, string sheetName, string rangeAddress, List<List<string>>? formats = null, string? formatsFile = null)
    {
        // Resolve formats from inline parameter or file
        var resolvedFormats = ParameterTransforms.ResolveFormulasOrFile(formats, formatsFile, "formats");

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
                if (resolvedFormats.Count != rowCount)
                {
                    throw new ArgumentException($"Format array row count ({resolvedFormats.Count}) doesn't match range row count ({rowCount})", nameof(formats));
                }

                for (int i = 0; i < resolvedFormats.Count; i++)
                {
                    if (resolvedFormats[i].Count != columnCount)
                    {
                        throw new ArgumentException($"Format array row {i + 1} column count ({resolvedFormats[i].Count}) doesn't match range column count ({columnCount})", nameof(formats));
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
                                ((Excel.Range)cell).NumberFormatLocal = ctx.FormatTranslator.TranslateToLocale(resolvedFormats[row - 1][col - 1]);
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
                            formatArray[row, col] = ctx.FormatTranslator.TranslateToLocale(resolvedFormats[row][col]);
                        }
                    }

                    // Set number formats via 2D array
                    ((Excel.Range)range).NumberFormatLocal = formatArray;
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
