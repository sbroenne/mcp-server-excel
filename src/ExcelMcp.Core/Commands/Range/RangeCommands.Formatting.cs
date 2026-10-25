using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Formatting operations for Excel ranges (partial class)
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public OperationResult SetStyle(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        string styleName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;

            try
            {
                sheet = Utilities.WorksheetLookup.Get(ctx.Book, sheetName);

                range = sheet.Range[rangeAddress];
                range.Style = styleName;

                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Action = "set-style"
                };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public RangeStyleResult GetStyle(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Styles? styles = null;
            Excel.Style? style = null;
            object? rangeStyle = null;

            try
            {
                sheet = Utilities.WorksheetLookup.Get(ctx.Book, sheetName);

                range = sheet.Range[rangeAddress];

                rangeStyle = range.Style;

                if (rangeStyle is null or DBNull)
                    rangeStyle = "Normal";
                if (rangeStyle is Excel.Style nativeStyle)
                {
                    style = nativeStyle;
                    rangeStyle = null;
                }
                else if (rangeStyle is string name)
                {
                    styles = ctx.Book.Styles;
                    style = styles[name];
                }
                else
                    throw new InvalidOperationException("Excel returned an unsupported cell-style identity. Use get-format to inspect each cell.");
                string styleName = style.Name;
                bool isBuiltIn = style.BuiltIn;
                string styleDescription = style.NameLocal;

                return new RangeStyleResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    SheetName = sheetName,
                    RangeAddress = range.Address,
                    StyleName = styleName,
                    IsBuiltInStyle = isBuiltIn,
                    StyleDescription = styleDescription
                };
            }
            finally
            {
                ComUtilities.Release(ref style!);
                ComUtilities.Release(ref styles!);
                ComUtilities.Release(ref rangeStyle);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    private static void ValidateRangeAddresses(string[] rangeAddresses)
    {
        if (rangeAddresses == null || rangeAddresses.Length == 0)
        {
            throw new ArgumentException("At least one range address is required.", nameof(rangeAddresses));
        }

        for (var index = 0; index < rangeAddresses.Length; index++)
        {
            if (string.IsNullOrWhiteSpace(rangeAddresses[index]))
            {
                throw new ArgumentException($"Range address at index {index} cannot be null, empty, or whitespace.", nameof(rangeAddresses));
            }
        }
    }

}
