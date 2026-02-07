using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;


namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range discovery operations (UsedRange, CurrentRegion, RangeInfo)
/// </summary>
public partial class RangeCommands
{
    // === NATIVE EXCEL COM OPERATIONS (AI/LLM ESSENTIAL) ===

    /// <summary>
    /// Gets the used range (all non-empty cells) from worksheet
    /// Excel COM: Worksheet.UsedRange
    /// </summary>
    public RangeValueResult GetUsedRange(IExcelBatch batch, string sheetName)
    {
        var result = new RangeValueResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                range = sheet.UsedRange;
                result.RangeAddress = range.Address;

                // Get values as 2D array
                object[,]? values = range.Value2;
                if (values != null)
                {
                    result.RowCount = values.GetLength(0);
                    result.ColumnCount = values.GetLength(1);

                    for (int r = 1; r <= result.RowCount; r++)
                    {
                        var row = new List<object?>();
                        for (int c = 1; c <= result.ColumnCount; c++)
                        {
                            row.Add(values[r, c]);
                        }
                        result.Values.Add(row);
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public RangeValueResult GetCurrentRegion(IExcelBatch batch, string sheetName, string cellAddress)
    {
        var result = new RangeValueResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = cellAddress
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? cell = null;
            dynamic? region = null;
            try
            {
                cell = RangeHelpers.ResolveRange(ctx.Book, sheetName, cellAddress, out string? specificError);
                if (cell == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, cellAddress));
                }

                region = cell.CurrentRegion;
                result.RangeAddress = region.Address;

                // Get values as 2D array
                object[,]? values = region.Value2;
                if (values != null)
                {
                    result.RowCount = values.GetLength(0);
                    result.ColumnCount = values.GetLength(1);

                    for (int r = 1; r <= result.RowCount; r++)
                    {
                        var row = new List<object?>();
                        for (int c = 1; c <= result.ColumnCount; c++)
                        {
                            row.Add(values[r, c]);
                        }
                        result.Values.Add(row);
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref region);
                ComUtilities.Release(ref cell);
            }
        });
    }

    /// <inheritdoc />
    public RangeInfoResult GetInfo(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new RangeInfoResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName
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

                result.Address = range.Address;
                result.RowCount = range.Rows.Count;
                result.ColumnCount = range.Columns.Count;
                result.NumberFormat = range.NumberFormat?.ToString();

                // Cell geometry properties (position and dimensions in points)
                result.Left = Convert.ToDouble(range.Left);
                result.Top = Convert.ToDouble(range.Top);
                result.Width = Convert.ToDouble(range.Width);
                result.Height = Convert.ToDouble(range.Height);

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



