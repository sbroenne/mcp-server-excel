using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

public partial class WorkbookCommands
{
    private static readonly Excel.XlTableStyleElementType[] TableStyleElementTypes =
        Enum.GetValues<Excel.XlTableStyleElementType>().Distinct().ToArray();

    /// <inheritdoc/>
    public TableStyleListResult ListTableStyles(IExcelBatch batch) =>
        batch.Execute((ctx, ct) =>
        {
            Excel.TableStyles? styles = null;
            try
            {
                styles = ctx.Book.TableStyles;
                var result = new TableStyleListResult { Success = true, FilePath = batch.WorkbookPath };
                for (int index = 1; index <= styles.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.TableStyle? style = null;
                    try
                    {
                        style = styles[index];
                        result.Styles.Add(ReadTableStyleInfo(style));
                    }
                    finally { ComUtilities.Release(ref style); }
                }
                return result;
            }
            finally { ComUtilities.Release(ref styles); }
        });

    /// <inheritdoc/>
    public TableStyleResult GetTableStyle(IExcelBatch batch, string styleName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.TableStyles? styles = null;
            Excel.TableStyle? style = null;
            try
            {
                styles = ctx.Book.TableStyles;
                style = FindTableStyle(styles, styleName, ct);
                return ReadTableStyleResult(style, batch.WorkbookPath, ct);
            }
            finally
            {
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref styles);
            }
        });
    }

    /// <inheritdoc/>
    public TableStyleResult CreateTableStyle(IExcelBatch batch, string styleName, string sourceStyleName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceStyleName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.TableStyles? styles = null;
            Excel.TableStyle? existing = null;
            Excel.TableStyle? source = null;
            Excel.TableStyle? style = null;
            try
            {
                styles = ctx.Book.TableStyles;
                existing = TryFindTableStyle(styles, styleName, ct);
                if (existing is not null)
                    throw new ArgumentException($"Table style '{styleName}' already exists.", nameof(styleName));
                source = FindTableStyle(styles, sourceStyleName, ct);
                ct.ThrowIfCancellationRequested();
                style = source.Duplicate(styleName);
                ComUtilities.Release(ref style);
                style = styles[styleName];
                return ReadTableStyleResult(style, batch.WorkbookPath, ct);
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
    public TableStyleResult UpdateTableStyle(IExcelBatch batch, string styleName, TableStyleOptions tableStyleOptions)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        ArgumentNullException.ThrowIfNull(tableStyleOptions);
        HashSet<Excel.XlTableStyleElementType> types = [];
        foreach (var setting in tableStyleOptions.Elements ?? [])
        {
            if (setting is null) throw new ArgumentException("Table style elements cannot be null.");
            var type = ParseTableStyleElement(setting.ElementType);
            if (!types.Add(type)) throw new ArgumentException("Table style elements must be unique, including native aliases.");
            CellFormatWriter.Validate(setting.ToFormatOptions());
            if (setting.StripeSize is < 1 || (setting.StripeSize is not null && !IsStripeElement(type)))
                throw new ArgumentException("stripeSize must be positive and is valid only for row/column stripe elements.");
            if (setting.Clear == true && (setting.HasFormatting || setting.StripeSize is not null))
                throw new ArgumentException("clear cannot be combined with formatting or stripeSize.");
            if ((setting.Borders ?? []).Any(border => border.Position is CellBorderPosition.DiagonalDown or CellBorderPosition.DiagonalUp))
                throw new ArgumentException("Native table-style elements do not support diagonal borders.");
        }
        return batch.Execute((ctx, ct) =>
        {
            Excel.TableStyles? styles = null;
            Excel.TableStyle? style = null;
            Excel.TableStyleElements? elements = null;
            try
            {
                styles = ctx.Book.TableStyles;
                style = FindTableStyle(styles, styleName, ct);
                RejectBuiltInTableStyle(style);
                elements = style.TableStyleElements;
                foreach (var setting in tableStyleOptions.Elements ?? [])
                {
                    ct.ThrowIfCancellationRequested();
                    if (setting.StripeSize is null || setting.HasFormatting) continue;
                    Excel.TableStyleElement? stripeElement = null;
                    try
                    {
                        stripeElement = elements.Item(ParseTableStyleElement(setting.ElementType));
                        if (!stripeElement.HasFormat)
                            throw new OperationFailureException(OperationFailureCategory.InvalidInput,
                                $"Element '{setting.ElementType}' has no native format. Supply formatting with stripeSize to create the stripe definition.");
                    }
                    finally { ComUtilities.Release(ref stripeElement); }
                }
                foreach (var setting in tableStyleOptions.Elements ?? [])
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.TableStyleElement? element = null;
                    Excel.Font? font = null;
                    Excel.Interior? fill = null;
                    Excel.Borders? borders = null;
                    try
                    {
                        element = elements.Item(ParseTableStyleElement(setting.ElementType));
                        if (setting.Clear == true) element.Clear();
                        else
                        {
                            var format = setting.ToFormatOptions();
                            // An unset element creates differential components lazily; acquire each immediately before writing.
                            if (setting.HasFillFormatting)
                            {
                                fill = element.Interior;
                                CellFormatWriter.ApplyFill(fill, format);
                            }
                            if (setting.HasFontFormatting)
                            {
                                font = element.Font;
                                CellFormatWriter.ApplyFont(font, format);
                            }
                            if (setting.Borders?.Count > 0)
                            {
                                borders = element.Borders;
                                CellFormatWriter.ApplyBorders(borders, format, ct);
                            }
                            if (setting.StripeSize is { } stripe) element.StripeSize = stripe;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref borders);
                        ComUtilities.Release(ref font);
                        ComUtilities.Release(ref fill);
                        ComUtilities.Release(ref element);
                    }
                }
                if (tableStyleOptions.ShowAsAvailableTableStyle is { } table) style.ShowAsAvailableTableStyle = table;
                if (tableStyleOptions.ShowAsAvailablePivotTableStyle is { } pivot) style.ShowAsAvailablePivotTableStyle = pivot;
                if (tableStyleOptions.ShowAsAvailableSlicerStyle is { } slicer) style.ShowAsAvailableSlicerStyle = slicer;
                if (tableStyleOptions.ShowAsAvailableTimelineStyle is { } timeline) style.ShowAsAvailableTimelineStyle = timeline;
                return ReadTableStyleResult(style, batch.WorkbookPath, ct);
            }
            finally
            {
                ComUtilities.Release(ref elements);
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref styles);
            }
        });
    }

    /// <inheritdoc/>
    public OperationResult DeleteTableStyle(IExcelBatch batch, string styleName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.TableStyles? styles = null;
            Excel.TableStyle? style = null;
            try
            {
                styles = ctx.Book.TableStyles;
                style = FindTableStyle(styles, styleName, ct);
                RejectBuiltInTableStyle(style);
                ct.ThrowIfCancellationRequested();
                style.Delete();
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath, Action = "delete-table-style" };
            }
            finally
            {
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref styles);
            }
        });
    }

    private static void RejectBuiltInTableStyle(Excel.TableStyle style)
    {
        if (style.BuiltIn) throw new InvalidOperationException("Built-in table styles are read-only. Clone one into a custom style first.");
    }

    private static Excel.XlTableStyleElementType ParseTableStyleElement(string value)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(value);
        if (!value.StartsWith("xl", StringComparison.OrdinalIgnoreCase) ||
            !Enum.TryParse<Excel.XlTableStyleElementType>(value, true, out var type) || !Enum.IsDefined(type))
            throw new ArgumentException($"Unknown native table-style element type: '{value}'.");
        return type;
    }

    private static bool IsStripeElement(Excel.XlTableStyleElementType type) =>
        type is Excel.XlTableStyleElementType.xlRowStripe1 or Excel.XlTableStyleElementType.xlRowStripe2 or
            Excel.XlTableStyleElementType.xlColumnStripe1 or Excel.XlTableStyleElementType.xlColumnStripe2;

    private static Excel.TableStyle FindTableStyle(Excel.TableStyles styles, string name, CancellationToken ct) =>
        TryFindTableStyle(styles, name, ct) ??
        throw new OperationFailureException(OperationFailureCategory.NotFound, $"Table style '{name}' was not found.");

    private static Excel.TableStyle? TryFindTableStyle(Excel.TableStyles styles, string name, CancellationToken ct)
    {
        for (int index = 1; index <= styles.Count; index++)
        {
            ct.ThrowIfCancellationRequested();
            Excel.TableStyle? style = null;
            try
            {
                style = styles[index];
                if (string.Equals(style.Name, name, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(style.NameLocal, name, StringComparison.OrdinalIgnoreCase))
                {
                    var result = style;
                    style = null;
                    return result;
                }
            }
            finally { ComUtilities.Release(ref style); }
        }
        return null;
    }

    private static TableStyleInfo ReadTableStyleInfo(Excel.TableStyle style) =>
        new(style.Name, style.NameLocal, style.BuiltIn, style.ShowAsAvailableTableStyle,
            style.ShowAsAvailablePivotTableStyle, style.ShowAsAvailableSlicerStyle, style.ShowAsAvailableTimelineStyle);

    private static TableStyleResult ReadTableStyleResult(Excel.TableStyle style, string filePath, CancellationToken ct)
    {
        Excel.TableStyleElements? elements = null;
        try
        {
            elements = style.TableStyleElements;
            if (elements.Count != TableStyleElementTypes.Length)
                throw new NotSupportedException($"Excel exposes {elements.Count} table-style elements, but the installed PIA describes {TableStyleElementTypes.Length}. A complete definition cannot be returned.");
            var info = ReadTableStyleInfo(style);
            List<TableStyleElementDefinition> definitions = [];
            foreach (var type in TableStyleElementTypes)
            {
                ct.ThrowIfCancellationRequested();
                Excel.TableStyleElement? element = null;
                Excel.Font? font = null;
                Excel.Interior? fill = null;
                Excel.Borders? borders = null;
                try
                {
                    element = elements.Item(type);
                    if (!element.HasFormat)
                    {
                        definitions.Add(new(type.ToString(), (int)type, false, null, null, null, null, []));
                        continue;
                    }
                    font = element.Font;
                    fill = element.Interior;
                    borders = element.Borders;
                    List<string> unset = [];
                    var visual = RangeCommands.ReadVisualFormat(font, fill, borders, unset, ct, tableStyleDefinition: true);
                    definitions.Add(new(type.ToString(), (int)type, true, IsStripeElement(type) ? element.StripeSize : null,
                        visual.Font, visual.Fill, visual.Borders, unset));
                }
                finally
                {
                    ComUtilities.Release(ref borders);
                    ComUtilities.Release(ref fill);
                    ComUtilities.Release(ref font);
                    ComUtilities.Release(ref element);
                }
            }
            return new TableStyleResult
            {
                Success = true,
                FilePath = filePath,
                Style = new(info.Name, info.NameLocal, info.BuiltIn, info.ShowAsAvailableTableStyle,
                    info.ShowAsAvailablePivotTableStyle, info.ShowAsAvailableSlicerStyle, info.ShowAsAvailableTimelineStyle, definitions)
            };
        }
        finally { ComUtilities.Release(ref elements); }
    }
}
