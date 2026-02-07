#pragma warning disable CS1591
namespace Sbroenne.ExcelMcp.Core.Models.Actions;

/// <summary>
/// Helper extensions to convert enum actions to string format and parse from cli strings.
/// </summary>
public static class ActionExtensions
{
    public static string ToActionString(this FileAction action) => action switch
    {
        FileAction.List => "list",
        FileAction.Open => "open",
        FileAction.Close => "close",
        FileAction.Create => "create",
        FileAction.CloseWorkbook => "close-workbook",
        FileAction.Test => "test",
        _ => throw new ArgumentException($"Unknown FileAction: {action}")
    };

    // NOTE: PowerQueryAction.ToActionString() is now generated in ServiceRegistry.PowerQuery.ToActionString()
    // NOTE: SheetAction.ToActionString() is now generated in ServiceRegistry.Sheet.ToActionString()
    // NOTE: SheetStyleAction.ToActionString() is now generated in ServiceRegistry.SheetStyle.ToActionString()
    // See Sbroenne.ExcelMcp.Generated namespace

    // NOTE: RangeAction.ToActionString() is now generated in ServiceRegistry.Range.ToActionString()
    // NOTE: RangeEditAction.ToActionString() is now generated in ServiceRegistry.RangeEdit.ToActionString()
    // NOTE: RangeFormatAction.ToActionString() is now generated in ServiceRegistry.RangeFormat.ToActionString()
    // NOTE: RangeLinkAction.ToActionString() is now generated in ServiceRegistry.RangeLink.ToActionString()

    // NamedRangeAction.ToActionString() is now generated in ServiceRegistry.NamedRange.ToActionString()

    // ConditionalFormatAction.ToActionString() is now generated in ServiceRegistry.ConditionalFormat.ToActionString()

    // VbaAction.ToActionString() is now generated in ServiceRegistry.Vba.ToActionString()

    // ConnectionAction.ToActionString() is now generated in ServiceRegistry.Connection.ToActionString()

    // DataModelAction.ToActionString() is now generated in ServiceRegistry.DataModel.ToActionString()

    // DataModelRelAction.ToActionString() is now generated in ServiceRegistry.DataModelRel.ToActionString()

    // TableAction.ToActionString() is now generated in ServiceRegistry.Table.ToActionString()

    // TableColumnAction.ToActionString() is now generated in ServiceRegistry.TableColumn.ToActionString()

    // PivotTableAction.ToActionString() is now generated in ServiceRegistry.PivotTable.ToActionString()
    // PivotTableFieldAction.ToActionString() is now generated in ServiceRegistry.PivotTableField.ToActionString()
    // PivotTableCalcAction.ToActionString() is now generated in ServiceRegistry.PivotTableCalc.ToActionString()

    // ChartAction.ToActionString() is now generated in ServiceRegistry.Chart.ToActionString()
    // ChartConfigAction.ToActionString() is now generated in ServiceRegistry.ChartConfig.ToActionString()

    // SlicerAction.ToActionString() is now generated in ServiceRegistry.Slicer.ToActionString()

    // CalculationModeAction.ToActionString() is now generated in ServiceRegistry.Calculation.ToActionString()
}
#pragma warning restore CS1591


