using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

public partial class PivotTableCommands
{
    private static string ReadSharingReportState(Excel.PivotTable pivot, string filePath, CancellationToken ct)
    {
        Excel.Range? extent = null;
        Excel.PivotFields? fields = null;
        Excel.PivotFields? rows = null;
        Excel.PivotField? valuesPlacement = null;
        string reading = "the report fields";
        try
        {
            extent = pivot.TableRange2;
            fields = (Excel.PivotFields)pivot.PivotFields();
            rows = (Excel.PivotFields)pivot.RowFields;
            List<object> settings = [];
            for (int i = 1; i <= fields.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.PivotField? field = null;
                Excel.PivotItems? items = null;
                try
                {
                    field = fields.Item(i);
                    reading = $"field '{field.Name}'";
                    if (field.Orientation == Excel.XlPivotFieldOrientation.xlHidden ||
                        PivotTableHelpers.IsValuesPlaceholderField(field, field.SourceName))
                        continue;
                    var selection = ReadRegularFieldSelection(field, pivot.Name, filePath, ct, int.MaxValue,
                        manualSelectionOnly: true);
                    if (selection is not null && (!selection.FieldIdentityAvailable || !selection.LayoutAvailable ||
                        !selection.SelectionStateAvailable || !selection.FiltersAvailable || selection.MembersTruncated))
                        throw new NotSupportedException(
                            $"Cannot verify field '{field.Name}' before sharing: {selection.IdentityError ?? selection.LayoutError ?? selection.SelectionError ?? selection.FiltersError}");
                    List<string> itemOrder = [];
                    bool axis = field.Orientation is Excel.XlPivotFieldOrientation.xlRowField or Excel.XlPivotFieldOrientation.xlColumnField;
                    int? sortOrder = axis ? field.AutoSortOrder : null;
                    if (sortOrder == (int)Excel.Constants.xlManual)
                    {
                        items = (Excel.PivotItems)field.PivotItems();
                        for (int j = 1; j <= items.Count; j++)
                        {
                            ct.ThrowIfCancellationRequested();
                            Excel.PivotItem? item = null;
                            try { item = items.Item(j); itemOrder.Add(item.Name); }
                            finally { ComUtilities.Release(ref item); }
                        }
                    }
                    settings.Add(new
                    {
                        field.Name,
                        field.SourceName,
                        field.Caption,
                        field.Orientation,
                        Position = Convert.ToInt32((object)field.Position, CultureInfo.InvariantCulture),
                        Selection = selection,
                        SortOrder = sortOrder,
                        SortField = axis ? field.AutoSortField : null,
                        ItemOrder = itemOrder,
                        Subtotals = axis ? ReadSubtotalFunctions(field, pivot.Name, filePath, ct) : null,
                        NumberFormat = field.Orientation == Excel.XlPivotFieldOrientation.xlDataField ? field.NumberFormat : null
                    });
                }
                finally
                {
                    ComUtilities.Release(ref items);
                    ComUtilities.Release(ref field);
                }
            }
            reading = "value calculations";
            var values = ReadAllValueFieldCalculations(pivot, filePath, ct);
            if (values.Count > 1)
                valuesPlacement = pivot.DataPivotField;
            reading = "row layout";
            var rowLayouts = ReadRowLayouts(rows, ct);
            reading = "report options";
            return JsonSerializer.Serialize(new
            {
                pivot.Name,
                extent.Row,
                extent.Column,
                Fields = settings,
                Values = values,
                ValuesPlacement = valuesPlacement is null ? null : new
                {
                    valuesPlacement.Orientation,
                    Position = Convert.ToInt32((object)valuesPlacement.Position, CultureInfo.InvariantCulture)
                },
                Style = ReadPivotStyleName(pivot),
                pivot.PreserveFormatting,
                pivot.ShowTableStyleRowStripes,
                pivot.ShowTableStyleColumnStripes,
                pivot.ShowTableStyleRowHeaders,
                pivot.ShowTableStyleColumnHeaders,
                pivot.RowGrand,
                pivot.ColumnGrand,
                pivot.AllowMultipleFilters,
                pivot.MergeLabels,
                pivot.DisplayNullString,
                pivot.NullString,
                pivot.DisplayErrorString,
                pivot.ErrorString,
                pivot.PrintTitles,
                pivot.RepeatItemsOnEachPrintedPage,
                RowFields = rowLayouts
            });
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"Excel could not read {reading} for cache-sharing verification.", ex);
        }
        finally
        {
            ComUtilities.Release(ref valuesPlacement);
            ComUtilities.Release(ref rows);
            ComUtilities.Release(ref fields);
            ComUtilities.Release(ref extent);
        }
    }
}
