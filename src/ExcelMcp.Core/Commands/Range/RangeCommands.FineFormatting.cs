using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    /// <inheritdoc/>
    public OperationResult Format(IExcelBatch batch, string sheetName,
        string[] rangeAddresses, CellFormatOptions formatOptions)
    {
        ValidateRangeAddresses(rangeAddresses);
        CellFormatWriter.Validate(formatOptions);
        return batch.Execute((ctx, ct) =>
        {
            List<Excel.Range> targets = [];
            try
            {
                string? localizedNumberFormat = formatOptions.NumberFormat is null
                    ? null : ctx.FormatTranslator.TranslateToLocale(formatOptions.NumberFormat);
                for (int index = 0; index < rangeAddresses.Length; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Range? range = null;
                    Excel.Worksheet? sheet = null;
                    Excel.Protection? protection = null;
                    try
                    {
                        try
                        {
                            range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddresses[index]);
                        }
                        catch (COMException ex)
                        {
                            throw new ArgumentException($"Invalid range address at index {index}: '{rangeAddresses[index]}'", nameof(rangeAddresses), ex);
                        }
                        catch (OperationFailureException ex) when (ex.ErrorCategory == OperationFailureCategory.InvalidInput)
                        {
                            throw new ArgumentException($"Invalid range address at index {index}: '{rangeAddresses[index]}'", nameof(rangeAddresses), ex);
                        }
                        sheet = (Excel.Worksheet)range.Parent;
                        if (sheet.ProtectContents)
                        {
                            protection = sheet.Protection;
                            if (!protection.AllowFormattingCells)
                                throw new InvalidOperationException($"Worksheet '{sheet.Name}' is protected against cell formatting.");
                        }
                        targets.Add(range);
                        range = null;
                    }
                    finally
                    {
                        ComUtilities.Release(ref protection);
                        ComUtilities.Release(ref sheet);
                        ComUtilities.Release(ref range);
                    }
                }
                foreach (var range in targets)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Font? font = null;
                    Excel.Interior? fill = null;
                    Excel.Borders? borders = null;
                    try
                    {
                        font = range.Font;
                        fill = range.Interior;
                        borders = range.Borders;
                        CellFormatWriter.ApplyFont(font, formatOptions);
                        CellFormatWriter.ApplyFill(fill, formatOptions);
                        CellFormatWriter.ApplyBorders(borders, formatOptions, ct);
                        if (formatOptions.HorizontalAlignment is not null)
                            range.HorizontalAlignment = CellFormatWriter.ParseHorizontalAlignment(formatOptions.HorizontalAlignment);
                        if (formatOptions.VerticalAlignment is not null)
                            range.VerticalAlignment = CellFormatWriter.ParseVerticalAlignment(formatOptions.VerticalAlignment);
                        if (formatOptions.WrapText is { } wrap) range.WrapText = wrap;
                        if (formatOptions.ShrinkToFit is { } shrink) range.ShrinkToFit = shrink;
                        if (formatOptions.IndentLevel is { } indent) range.IndentLevel = indent;
                        if (formatOptions.ReadingOrder is not null)
                            range.ReadingOrder = CellFormatWriter.ParseReadingOrder(formatOptions.ReadingOrder);
                        if (formatOptions.Orientation is { } orientation) range.Orientation = orientation;
                        if (localizedNumberFormat is not null) range.NumberFormatLocal = localizedNumberFormat;
                    }
                    finally
                    {
                        ComUtilities.Release(ref borders);
                        ComUtilities.Release(ref fill);
                        ComUtilities.Release(ref font);
                    }
                }
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath, Action = "format" };
            }
            finally
            {
                for (int index = targets.Count - 1; index >= 0; index--)
                {
                    var range = targets[index];
                    ComUtilities.Release(ref range);
                }
            }
        });
    }
}
