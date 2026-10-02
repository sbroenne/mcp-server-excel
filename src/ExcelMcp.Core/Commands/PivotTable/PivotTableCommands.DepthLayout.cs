using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

public partial class PivotTableCommands
{
    /// <inheritdoc/>
    public PivotLayoutResult GetLayoutOptions(IExcelBatch batch, string pivotTableName) =>
        PivotLayout(batch, pivotTableName, null);

    /// <inheritdoc/>
    public PivotLayoutResult SetLayoutOptions(IExcelBatch batch, string pivotTableName, PivotLayoutOptions layoutOptions)
    {
        ArgumentNullException.ThrowIfNull(layoutOptions);
        if (layoutOptions.RowLayout is < 0 or > 2)
            throw new ArgumentOutOfRangeException(nameof(layoutOptions), "rowLayout must be 0, 1, or 2.");
        if (layoutOptions.StyleName is not null)
            ArgumentException.ThrowIfNullOrWhiteSpace(layoutOptions.StyleName);
        if (layoutOptions.RepeatLabels == true && layoutOptions.RowLayout == 0)
            throw new ArgumentException("Repeated labels require Tabular or Outline layout.");
        return PivotLayout(batch, pivotTableName, layoutOptions);
    }

    private static PivotLayoutResult PivotLayout(IExcelBatch batch, string name, PivotLayoutOptions? options)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.Worksheet? sheet = null;
            Excel.TableStyles? styles = null;
            Excel.TableStyle? style = null;
            Excel.PivotFields? rows = null;
            try
            {
                pivot = (Excel.PivotTable)FindPivotTable(ctx.Book, name);
                sheet = (Excel.Worksheet)pivot.Parent;
                if (options is not null && sheet.ProtectContents)
                    throw new InvalidOperationException("Unprotect the PivotTable worksheet before changing its layout.");
                rows = (Excel.PivotFields)pivot.RowFields;
                if (options?.StyleName is not null)
                {
                    styles = ctx.Book.TableStyles;
                    style = styles[options.StyleName];
                    if (!style.ShowAsAvailablePivotTableStyle)
                        throw new ArgumentException("The selected style is not available for PivotTables.");
                }
                if (options?.RepeatLabels == true && options.RowLayout is null)
                {
                    foreach (var field in ReadRowLayouts(rows, ct))
                        if (field.RowLayout == 0)
                            throw new ArgumentException("Repeated labels require a noncompact layout; specify rowLayout=1 or 2.");
                }
                ct.ThrowIfCancellationRequested();
                if (options is not null)
                {
                    if (options.RowLayout.HasValue)
                        pivot.RowAxisLayout(options.RowLayout.Value switch
                        {
                            0 => Excel.XlLayoutRowType.xlCompactRow,
                            1 => Excel.XlLayoutRowType.xlTabularRow,
                            _ => Excel.XlLayoutRowType.xlOutlineRow
                        });
                    if (options.RepeatLabels.HasValue)
                        pivot.RepeatAllLabels(options.RepeatLabels.Value ? Excel.XlPivotFieldRepeatLabels.xlRepeatLabels : Excel.XlPivotFieldRepeatLabels.xlDoNotRepeatLabels);
                    if (style is not null)
                        pivot.TableStyle2 = style.Name;
                    if (options.PreserveFormatting.HasValue)
                        pivot.PreserveFormatting = options.PreserveFormatting.Value;
                    if (options.ShowRowStripes.HasValue)
                        pivot.ShowTableStyleRowStripes = options.ShowRowStripes.Value;
                    if (options.ShowColumnStripes.HasValue)
                        pivot.ShowTableStyleColumnStripes = options.ShowColumnStripes.Value;
                    if (options.ShowRowHeaders.HasValue)
                        pivot.ShowTableStyleRowHeaders = options.ShowRowHeaders.Value;
                    if (options.ShowColumnHeaders.HasValue)
                        pivot.ShowTableStyleColumnHeaders = options.ShowColumnHeaders.Value;
                    if (options.AllowMultipleFilters.HasValue)
                        pivot.AllowMultipleFilters = options.AllowMultipleFilters.Value;
                }
                return new PivotLayoutResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    PivotTableName = name,
                    StyleName = ReadPivotStyleName(pivot),
                    PreserveFormatting = pivot.PreserveFormatting,
                    ShowRowStripes = pivot.ShowTableStyleRowStripes,
                    ShowColumnStripes = pivot.ShowTableStyleColumnStripes,
                    ShowRowHeaders = pivot.ShowTableStyleRowHeaders,
                    ShowColumnHeaders = pivot.ShowTableStyleColumnHeaders,
                    AllowMultipleFilters = pivot.AllowMultipleFilters,
                    RowFields = ReadRowLayouts(rows, ct)
                };
            }
            finally
            {
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref styles);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    private static string ReadPivotStyleName(Excel.PivotTable pivot)
    {
        object? value = null;
        try
        {
            value = pivot.TableStyle2;
            return value switch
            {
                Excel.TableStyle style => style.Name,
                string name => name,
                null => string.Empty,
                _ => throw new InvalidOperationException("Excel returned an unsupported PivotTable style value.")
            };
        }
        finally
        {
            ComUtilities.Release(ref value);
        }
    }

    private static List<PivotRowLayoutInfo> ReadRowLayouts(Excel.PivotFields rows, CancellationToken ct)
    {
        List<PivotRowLayoutInfo> result = [];
        for (int index = 1; index <= rows.Count; index++)
        {
            ct.ThrowIfCancellationRequested();
            Excel.PivotField? field = null;
            try
            {
                field = rows.Item(index);
                result.Add(new PivotRowLayoutInfo
                {
                    FieldName = field.Name,
                    RowLayout = field.LayoutCompactRow ? 0 : field.LayoutForm == Excel.XlLayoutFormType.xlTabular ? 1 : 2,
                    RepeatLabels = field.RepeatLabels
                });
            }
            finally
            {
                ComUtilities.Release(ref field);
            }
        }
        return result;
    }
}
