using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;


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
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new OperationFailureException(
                        OperationFailureCategory.NotFound,
                        $"Sheet '{sheetName}' not found.");
                }

                range = sheet.UsedRange;
                result.RangeAddress = range.Address;

                var values = ExcelValueNormalizer.Normalize((object?)range.Value2);
                result.RowCount = values.RowCount;
                result.ColumnCount = values.ColumnCount;
                result.Values.AddRange(values.Values);

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
            Excel.Range? cell = null;
            Excel.Range? region = null;
            try
            {
                cell = RangeHelpers.ResolveRange(ctx.Book, sheetName, cellAddress, out string? specificError);
                if (cell == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, cellAddress));
                }

                region = cell.CurrentRegion;
                result.RangeAddress = region.Address;

                var values = ExcelValueNormalizer.Normalize((object?)region.Value2);
                result.RowCount = values.RowCount;
                result.ColumnCount = values.ColumnCount;
                result.Values.AddRange(values.Values);

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
            Excel.Range? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                result.Address = range.Address;
                (result.RowCount, result.ColumnCount) = RangeHelpers.GetDimensions(range);
                object? numberFormat = range.NumberFormat;
                result.NumberFormat = numberFormat?.ToString();

                // Cell geometry properties (position and dimensions in points)
                result.Left = Convert.ToDouble((object)range.Left, System.Globalization.CultureInfo.InvariantCulture);
                result.Top = Convert.ToDouble((object)range.Top, System.Globalization.CultureInfo.InvariantCulture);
                result.Width = Convert.ToDouble((object)range.Width, System.Globalization.CultureInfo.InvariantCulture);
                result.Height = Convert.ToDouble((object)range.Height, System.Globalization.CultureInfo.InvariantCulture);

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
