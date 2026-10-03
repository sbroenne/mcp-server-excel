using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

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
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

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
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? styles = null;
            dynamic? style = null;
            object? rangeStyle = null;

            try
            {
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                range = sheet.Range[rangeAddress];

                try
                {
                    rangeStyle = range.Style;
                }
                catch (COMException)
                {
                    rangeStyle = "Normal";
                }

                if (rangeStyle is null or DBNull)
                    rangeStyle = "Normal";
                if (rangeStyle is Microsoft.Office.Interop.Excel.Style)
                {
                    style = rangeStyle;
                    rangeStyle = null;
                }
                else if (rangeStyle is string name)
                {
                    styles = ctx.Book.Styles;
                    style = styles.Item(name);
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
