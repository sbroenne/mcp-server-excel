using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

public partial class WorkbookCommands
{
    /// <inheritdoc/>
    public CellStyleListResult ListCellStyles(IExcelBatch batch) =>
        batch.Execute((ctx, ct) =>
        {
            Excel.Styles? styles = null;
            try
            {
                styles = ctx.Book.Styles;
                var result = new CellStyleListResult { Success = true, FilePath = batch.WorkbookPath };
                for (int index = 1; index <= styles.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Style? style = null;
                    try
                    {
                        style = styles[index];
                        result.Styles.Add(new CellStyleInfo(style.Name, style.NameLocal, style.BuiltIn));
                    }
                    finally
                    {
                        ComUtilities.Release(ref style);
                    }
                }
                return result;
            }
            finally
            {
                ComUtilities.Release(ref styles);
            }
        });

    /// <inheritdoc/>
    public CellStyleResult GetCellStyle(IExcelBatch batch, string styleName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Styles? styles = null;
            Excel.Style? style = null;
            try
            {
                styles = ctx.Book.Styles;
                style = FindCellStyle(styles, styleName, ct);
                return ReadCellStyleResult(style, batch.WorkbookPath, ct);
            }
            finally
            {
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref styles);
            }
        });
    }

    /// <inheritdoc/>
    public CellStyleResult CreateCellStyle(IExcelBatch batch, string styleName,
        string sourceSheetName, string sourceCellAddress)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceSheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceCellAddress);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Styles? styles = null;
            Excel.Style? existing = null;
            Excel.Style? style = null;
            Excel.Range? source = null;
            try
            {
                styles = ctx.Book.Styles;
                existing = TryFindCellStyle(styles, styleName, ct);
                if (existing is not null)
                    throw new ArgumentException($"Cell style '{styleName}' already exists.", nameof(styleName));
                source = RangeHelpers.ResolveRange(ctx.Book, sourceSheetName, sourceCellAddress);
                if (Convert.ToDouble(source.CountLarge) != 1)
                    throw new ArgumentException("sourceCellAddress must select exactly one cell.", nameof(sourceCellAddress));
                ct.ThrowIfCancellationRequested();
                style = AddCellStyleFromSource(ctx.App, styles, styleName, source);
                // Inspect the registered definition: the Add return can expose incomplete source-cell getters.
                ComUtilities.Release(ref style);
                style = styles[styleName];
                return ReadCellStyleResult(style, batch.WorkbookPath, ct);
            }
            finally
            {
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref source);
                ComUtilities.Release(ref existing);
                ComUtilities.Release(ref styles);
            }
        });
    }

    /// <inheritdoc/>
    public CellStyleResult UpdateCellStyle(IExcelBatch batch, string styleName, CellStyleOptions styleOptions)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        ArgumentNullException.ThrowIfNull(styleOptions);
        if (styleOptions.FormatOptions is not null)
        {
            CellFormatWriter.Validate(styleOptions.FormatOptions);
            foreach (var border in styleOptions.FormatOptions.Borders ?? [])
                _ = CellFormatWriter.StyleBorderIndex(border.Position);
        }
        return batch.Execute((ctx, ct) =>
        {
            Excel.Styles? styles = null;
            Excel.Style? style = null;
            Excel.Font? font = null;
            Excel.Interior? fill = null;
            Excel.Borders? borders = null;
            try
            {
                styles = ctx.Book.Styles;
                style = FindCellStyle(styles, styleName, ct);
                RejectBuiltInStyle(style);
                ct.ThrowIfCancellationRequested();
                var originalFlags = (style.IncludeFont, style.IncludeNumber, style.IncludeAlignment,
                    style.IncludeBorder, style.IncludePatterns, style.IncludeProtection);
                if (styleOptions.FormatOptions is { } options)
                {
                    string? localFormat = options.NumberFormat is null
                        ? null : ctx.FormatTranslator.TranslateToLocale(options.NumberFormat);
                    font = style.Font;
                    fill = style.Interior;
                    borders = style.Borders;
                    CellFormatWriter.ApplyFont(font, options);
                    CellFormatWriter.ApplyFill(fill, options);
                    CellFormatWriter.ApplyBorders(borders, options, ct, styleDefinition: true);
                    if (options.HorizontalAlignment is not null)
                        style.HorizontalAlignment = (Excel.XlHAlign)CellFormatWriter.ParseHorizontalAlignment(options.HorizontalAlignment);
                    if (options.VerticalAlignment is not null)
                        style.VerticalAlignment = (Excel.XlVAlign)CellFormatWriter.ParseVerticalAlignment(options.VerticalAlignment);
                    if (options.WrapText is { } wrap) style.WrapText = wrap;
                    if (options.ShrinkToFit is { } shrink) style.ShrinkToFit = shrink;
                    if (options.IndentLevel is { } indent) style.IndentLevel = indent;
                    if (options.ReadingOrder is not null) style.ReadingOrder = CellFormatWriter.ParseReadingOrder(options.ReadingOrder);
                    if (options.Orientation is { } orientation) style.Orientation = (Excel.XlOrientation)orientation;
                    if (localFormat is not null) style.NumberFormatLocal = localFormat;
                }
                if (styleOptions.Locked is { } locked) style.Locked = locked;
                if (styleOptions.FormulaHidden is { } hidden) style.FormulaHidden = hidden;
                // Native definition writes can enable components; preserve omitted inclusion choices.
                style.IncludeFont = styleOptions.IncludeFont ?? originalFlags.IncludeFont;
                style.IncludeNumber = styleOptions.IncludeNumber ?? originalFlags.IncludeNumber;
                style.IncludeAlignment = styleOptions.IncludeAlignment ?? originalFlags.IncludeAlignment;
                style.IncludeBorder = styleOptions.IncludeBorder ?? originalFlags.IncludeBorder;
                style.IncludePatterns = styleOptions.IncludePatterns ?? originalFlags.IncludePatterns;
                style.IncludeProtection = styleOptions.IncludeProtection ?? originalFlags.IncludeProtection;
                return ReadCellStyleResult(style, batch.WorkbookPath, ct);
            }
            finally
            {
                ComUtilities.Release(ref borders);
                ComUtilities.Release(ref fill);
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref styles);
            }
        });
    }

    /// <inheritdoc/>
    public OperationResult DeleteCellStyle(IExcelBatch batch, string styleName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Styles? styles = null;
            Excel.Style? style = null;
            try
            {
                styles = ctx.Book.Styles;
                style = FindCellStyle(styles, styleName, ct);
                RejectBuiltInStyle(style);
                ct.ThrowIfCancellationRequested();
                style.Delete();
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath, Action = "delete-cell-style" };
            }
            finally
            {
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref styles);
            }
        });
    }

    private static void RejectBuiltInStyle(Excel.Style style)
    {
        if (style.BuiltIn)
            throw new InvalidOperationException("Built-in cell styles are read-only. Create a custom style from one source cell instead.");
    }

    private static Excel.Style AddCellStyleFromSource(Excel.Application app, Excel.Styles styles,
        string styleName, Excel.Range source)
    {
        Excel.Worksheet? sourceSheet = null;
        Excel.Window? originalWindow = null;
        object? originalSheet = null;
        Excel.Style? created = null;
        Exception? operationError = null;
        bool activated = false;
        try
        {
            try
            {
                sourceSheet = (Excel.Worksheet)source.Parent;
                if (sourceSheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                    throw new OperationFailureException(OperationFailureCategory.InvalidInput,
                        "The source worksheet must be visible to capture a native cell style. Its visibility is not changed automatically.");
                originalWindow = app.ActiveWindow;
                originalSheet = app.ActiveSheet;
                if (originalWindow is null || originalSheet is not (Excel.Worksheet or Excel.Chart))
                    throw new InvalidOperationException("The original Excel view cannot be restored.");
                // Styles.Add requires the source sheet to be active, even with an explicit BasedOn range.
                activated = true;
                sourceSheet.Activate();
                created = styles.Add(styleName, source);
            }
            catch (Exception ex)
            {
                operationError = ex;
            }
            Exception? restoreError = null;
            try
            {
                if (activated)
                {
                    originalWindow!.Activate();
                    if (originalSheet is Excel.Worksheet sheet) sheet.Activate();
                    else if (originalSheet is Excel.Chart chart) chart.Activate();
                }
            }
            catch (Exception ex)
            {
                restoreError = ex;
            }
            if (restoreError is not null)
            {
                if (operationError is not null)
                    throw new AggregateException("Cell-style capture and restoring the original Excel view both failed.", operationError, restoreError);
                throw new InvalidOperationException("Cell-style capture completed, but restoring the original Excel view failed. Inspect the workbook before retrying.", restoreError);
            }
            if (operationError is not null)
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(operationError).Throw();
            var result = created ?? throw new InvalidOperationException("Excel did not return the created cell style.");
            created = null;
            return result;
        }
        finally
        {
            ComUtilities.Release(ref created);
            ComUtilities.Release(ref originalSheet);
            ComUtilities.Release(ref originalWindow);
            ComUtilities.Release(ref sourceSheet);
        }
    }

    private static Excel.Style FindCellStyle(Excel.Styles styles, string styleName, CancellationToken ct) =>
        TryFindCellStyle(styles, styleName, ct) ??
        throw new OperationFailureException(OperationFailureCategory.NotFound, $"Cell style '{styleName}' was not found.");

    private static Excel.Style? TryFindCellStyle(Excel.Styles styles, string styleName, CancellationToken ct)
    {
        for (int index = 1; index <= styles.Count; index++)
        {
            ct.ThrowIfCancellationRequested();
            Excel.Style? style = null;
            try
            {
                style = styles[index];
                if (string.Equals(style.Name, styleName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(style.NameLocal, styleName, StringComparison.OrdinalIgnoreCase))
                {
                    var selected = style;
                    style = null;
                    return selected;
                }
            }
            finally
            {
                ComUtilities.Release(ref style);
            }
        }
        return null;
    }

    private static CellStyleResult ReadCellStyleResult(Excel.Style style, string filePath, CancellationToken ct)
    {
        var format = RangeCommands.ReadStyleFormat(style, ct);
        return new()
        {
            Success = true,
            FilePath = filePath,
            Style = new CellStyleDefinition(style.Name, style.NameLocal, style.BuiltIn,
                style.IncludeFont, style.IncludeNumber, style.IncludeAlignment,
                style.IncludeBorder, style.IncludePatterns, style.IncludeProtection,
                format),
            ReadLimitations = ["Cell styles have no column-width/row-height settings or inside borders. Excel does not expose Style.MergeCells for inspection.",
                .. format.MixedFields.Where(field => field.EndsWith(".unavailable", StringComparison.Ordinal))
                    .Select(field => $"Excel does not expose {field[..^12]} for this partial cell-style definition.")]
        };
    }
}
