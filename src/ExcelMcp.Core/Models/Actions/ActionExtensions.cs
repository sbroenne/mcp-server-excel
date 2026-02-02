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

    public static string ToActionString(this PowerQueryAction action) => action switch
    {
        PowerQueryAction.List => "list",
        PowerQueryAction.View => "view",
        PowerQueryAction.Refresh => "refresh",
        PowerQueryAction.Delete => "delete",
        PowerQueryAction.GetLoadConfig => "get-load-config",
        PowerQueryAction.Create => "create",
        PowerQueryAction.Update => "update",
        PowerQueryAction.Rename => "rename",
        PowerQueryAction.RefreshAll => "refresh-all",
        PowerQueryAction.LoadTo => "load-to",
        PowerQueryAction.Unload => "unload",
        PowerQueryAction.Evaluate => "evaluate",
        _ => throw new ArgumentException($"Unknown PowerQueryAction: {action}")
    };

    public static string ToActionString(this WorksheetAction action) => action switch
    {
        WorksheetAction.List => "list",
        WorksheetAction.Create => "create",
        WorksheetAction.Rename => "rename",
        WorksheetAction.Copy => "copy",
        WorksheetAction.Delete => "delete",
        WorksheetAction.Move => "move",
        WorksheetAction.CopyToFile => "copy-to-file",
        WorksheetAction.MoveToFile => "move-to-file",
        _ => throw new ArgumentException($"Unknown WorksheetAction: {action}")
    };

    public static string ToActionString(this WorksheetStyleAction action) => action switch
    {
        WorksheetStyleAction.SetTabColor => "set-tab-color",
        WorksheetStyleAction.GetTabColor => "get-tab-color",
        WorksheetStyleAction.ClearTabColor => "clear-tab-color",
        WorksheetStyleAction.Hide => "hide",
        WorksheetStyleAction.VeryHide => "very-hide",
        WorksheetStyleAction.Show => "show",
        WorksheetStyleAction.GetVisibility => "get-visibility",
        WorksheetStyleAction.SetVisibility => "set-visibility",
        _ => throw new ArgumentException($"Unknown WorksheetStyleAction: {action}")
    };

    public static string ToActionString(this RangeAction action) => action switch
    {
        RangeAction.GetValues => "get-values",
        RangeAction.SetValues => "set-values",
        RangeAction.GetFormulas => "get-formulas",
        RangeAction.SetFormulas => "set-formulas",
        RangeAction.GetNumberFormats => "get-number-formats",
        RangeAction.SetNumberFormat => "set-number-format",
        RangeAction.SetNumberFormats => "set-number-formats",
        RangeAction.ClearAll => "clear-all",
        RangeAction.ClearContents => "clear-contents",
        RangeAction.ClearFormats => "clear-formats",
        RangeAction.Copy => "copy",
        RangeAction.CopyValues => "copy-values",
        RangeAction.CopyFormulas => "copy-formulas",
        RangeAction.GetUsedRange => "get-used-range",
        RangeAction.GetCurrentRegion => "get-current-region",
        RangeAction.GetInfo => "get-info",
        _ => throw new ArgumentException($"Unknown RangeAction: {action}")
    };

    public static string ToActionString(this RangeEditAction action) => action switch
    {
        RangeEditAction.InsertCells => "insert-cells",
        RangeEditAction.DeleteCells => "delete-cells",
        RangeEditAction.InsertRows => "insert-rows",
        RangeEditAction.DeleteRows => "delete-rows",
        RangeEditAction.InsertColumns => "insert-columns",
        RangeEditAction.DeleteColumns => "delete-columns",
        RangeEditAction.Find => "find",
        RangeEditAction.Replace => "replace",
        RangeEditAction.Sort => "sort",
        _ => throw new ArgumentException($"Unknown RangeEditAction: {action}")
    };

    public static string ToActionString(this RangeFormatAction action) => action switch
    {
        RangeFormatAction.GetStyle => "get-style",
        RangeFormatAction.SetStyle => "set-style",
        RangeFormatAction.FormatRange => "format-range",
        RangeFormatAction.ValidateRange => "validate-range",
        RangeFormatAction.GetValidation => "get-validation",
        RangeFormatAction.RemoveValidation => "remove-validation",
        RangeFormatAction.AutoFitColumns => "auto-fit-columns",
        RangeFormatAction.AutoFitRows => "auto-fit-rows",
        RangeFormatAction.MergeCells => "merge-cells",
        RangeFormatAction.UnmergeCells => "unmerge-cells",
        RangeFormatAction.GetMergeInfo => "get-merge-info",
        _ => throw new ArgumentException($"Unknown RangeFormatAction: {action}")
    };

    public static string ToActionString(this RangeLinkAction action) => action switch
    {
        RangeLinkAction.AddHyperlink => "add-hyperlink",
        RangeLinkAction.RemoveHyperlink => "remove-hyperlink",
        RangeLinkAction.ListHyperlinks => "list-hyperlinks",
        RangeLinkAction.GetHyperlink => "get-hyperlink",
        RangeLinkAction.SetCellLock => "set-cell-lock",
        RangeLinkAction.GetCellLock => "get-cell-lock",
        _ => throw new ArgumentException($"Unknown RangeLinkAction: {action}")
    };

    public static string ToActionString(this NamedRangeAction action) => action switch
    {
        NamedRangeAction.List => "list",
        NamedRangeAction.Read => "read",
        NamedRangeAction.Write => "write",
        NamedRangeAction.Create => "create",
        NamedRangeAction.Update => "update",
        NamedRangeAction.Delete => "delete",
        _ => throw new ArgumentException($"Unknown NamedRangeAction: {action}")
    };

    public static string ToActionString(this ConditionalFormatAction action) => action switch
    {
        ConditionalFormatAction.AddRule => "add-rule",
        ConditionalFormatAction.ClearRules => "clear-rules",
        _ => throw new ArgumentException($"Unknown ConditionalFormatAction: {action}")
    };

    public static string ToActionString(this VbaAction action) => action switch
    {
        VbaAction.List => "list",
        VbaAction.View => "view",
        VbaAction.Import => "import",
        VbaAction.Delete => "delete",
        VbaAction.Run => "run",
        VbaAction.Update => "update",
        _ => throw new ArgumentException($"Unknown VbaAction: {action}")
    };

    public static string ToActionString(this ConnectionAction action) => action switch
    {
        ConnectionAction.List => "list",
        ConnectionAction.View => "view",
        ConnectionAction.Create => "create",
        ConnectionAction.Test => "test",
        ConnectionAction.Refresh => "refresh",
        ConnectionAction.Delete => "delete",
        ConnectionAction.LoadTo => "load-to",
        ConnectionAction.GetProperties => "get-properties",
        ConnectionAction.SetProperties => "set-properties",
        _ => throw new ArgumentException($"Unknown ConnectionAction: {action}")
    };

    public static string ToActionString(this DataModelAction action) => action switch
    {
        DataModelAction.ListTables => "list-tables",
        DataModelAction.ReadTable => "read-table",
        DataModelAction.ListColumns => "list-columns",
        DataModelAction.ListMeasures => "list-measures",
        DataModelAction.Read => "read",
        DataModelAction.CreateMeasure => "create-measure",
        DataModelAction.UpdateMeasure => "update-measure",
        DataModelAction.DeleteMeasure => "delete-measure",
        DataModelAction.RenameTable => "rename-table",
        DataModelAction.DeleteTable => "delete-table",
        DataModelAction.ReadInfo => "read-info",
        DataModelAction.Refresh => "refresh",
        DataModelAction.Evaluate => "evaluate",
        DataModelAction.ExecuteDmv => "execute-dmv",
        _ => throw new ArgumentException($"Unknown DataModelAction: {action}")
    };

    public static string ToActionString(this DataModelRelAction action) => action switch
    {
        DataModelRelAction.ListRelationships => "list-relationships",
        DataModelRelAction.ReadRelationship => "read-relationship",
        DataModelRelAction.CreateRelationship => "create-relationship",
        DataModelRelAction.UpdateRelationship => "update-relationship",
        DataModelRelAction.DeleteRelationship => "delete-relationship",
        _ => throw new ArgumentException($"Unknown DataModelRelAction: {action}")
    };

    public static string ToActionString(this TableAction action) => action switch
    {
        TableAction.List => "list",
        TableAction.Read => "read",
        TableAction.Create => "create",
        TableAction.Rename => "rename",
        TableAction.Delete => "delete",
        TableAction.Resize => "resize",
        TableAction.SetStyle => "set-style",
        TableAction.ToggleTotals => "toggle-totals",
        TableAction.SetColumnTotal => "set-column-total",
        TableAction.Append => "append",
        TableAction.GetData => "get-data",
        TableAction.AddToDataModel => "add-to-datamodel",
        TableAction.CreateFromDax => "create-from-dax",
        TableAction.UpdateDax => "update-dax",
        TableAction.GetDax => "get-dax",
        _ => throw new ArgumentException($"Unknown TableAction: {action}")
    };

    public static string ToActionString(this TableColumnAction action) => action switch
    {
        TableColumnAction.ApplyFilter => "apply-filter",
        TableColumnAction.ApplyFilterValues => "apply-filter-values",
        TableColumnAction.ClearFilters => "clear-filters",
        TableColumnAction.GetFilters => "get-filters",
        TableColumnAction.AddColumn => "add-column",
        TableColumnAction.RemoveColumn => "remove-column",
        TableColumnAction.RenameColumn => "rename-column",
        TableColumnAction.GetStructuredReference => "get-structured-reference",
        TableColumnAction.Sort => "sort",
        TableColumnAction.SortMulti => "sort-multi",
        TableColumnAction.GetColumnNumberFormat => "get-column-number-format",
        TableColumnAction.SetColumnNumberFormat => "set-column-number-format",
        _ => throw new ArgumentException($"Unknown TableColumnAction: {action}")
    };

    public static string ToActionString(this PivotTableAction action) => action switch
    {
        PivotTableAction.List => "list",
        PivotTableAction.Read => "read",
        PivotTableAction.CreateFromRange => "create-from-range",
        PivotTableAction.CreateFromTable => "create-from-table",
        PivotTableAction.CreateFromDataModel => "create-from-datamodel",
        PivotTableAction.Delete => "delete",
        PivotTableAction.Refresh => "refresh",
        _ => throw new ArgumentException($"Unknown PivotTableAction: {action}")
    };

    public static string ToActionString(this PivotTableFieldAction action) => action switch
    {
        PivotTableFieldAction.ListFields => "list-fields",
        PivotTableFieldAction.AddRowField => "add-row-field",
        PivotTableFieldAction.AddColumnField => "add-column-field",
        PivotTableFieldAction.AddValueField => "add-value-field",
        PivotTableFieldAction.AddFilterField => "add-filter-field",
        PivotTableFieldAction.RemoveField => "remove-field",
        PivotTableFieldAction.SetFieldFunction => "set-field-function",
        PivotTableFieldAction.SetFieldName => "set-field-name",
        PivotTableFieldAction.SetFieldFormat => "set-field-format",
        PivotTableFieldAction.SetFieldFilter => "set-field-filter",
        PivotTableFieldAction.SortField => "sort-field",
        PivotTableFieldAction.GroupByDate => "group-by-date",
        PivotTableFieldAction.GroupByNumeric => "group-by-numeric",
        _ => throw new ArgumentException($"Unknown PivotTableFieldAction: {action}")
    };

    public static string ToActionString(this PivotTableCalcAction action) => action switch
    {
        PivotTableCalcAction.ListCalculatedFields => "list-calculated-fields",
        PivotTableCalcAction.CreateCalculatedField => "create-calculated-field",
        PivotTableCalcAction.DeleteCalculatedField => "delete-calculated-field",
        PivotTableCalcAction.ListCalculatedMembers => "list-calculated-members",
        PivotTableCalcAction.CreateCalculatedMember => "create-calculated-member",
        PivotTableCalcAction.DeleteCalculatedMember => "delete-calculated-member",
        PivotTableCalcAction.SetLayout => "set-layout",
        PivotTableCalcAction.SetSubtotals => "set-subtotals",
        PivotTableCalcAction.SetGrandTotals => "set-grand-totals",
        PivotTableCalcAction.GetData => "get-data",
        _ => throw new ArgumentException($"Unknown PivotTableCalcAction: {action}")
    };

    public static string ToActionString(this ChartAction action) => action switch
    {
        ChartAction.List => "list",
        ChartAction.Read => "read",
        ChartAction.CreateFromRange => "create-from-range",
        ChartAction.CreateFromTable => "create-from-table",
        ChartAction.CreateFromPivotTable => "create-from-pivottable",
        ChartAction.Delete => "delete",
        ChartAction.Move => "move",
        ChartAction.FitToRange => "fit-to-range",
        _ => throw new ArgumentException($"Unknown ChartAction: {action}")
    };

    public static string ToActionString(this ChartConfigAction action) => action switch
    {
        ChartConfigAction.SetSourceRange => "set-source-range",
        ChartConfigAction.AddSeries => "add-series",
        ChartConfigAction.RemoveSeries => "remove-series",
        ChartConfigAction.SetChartType => "set-chart-type",
        ChartConfigAction.SetTitle => "set-title",
        ChartConfigAction.SetAxisTitle => "set-axis-title",
        ChartConfigAction.GetAxisNumberFormat => "get-axis-number-format",
        ChartConfigAction.SetAxisNumberFormat => "set-axis-number-format",
        ChartConfigAction.ShowLegend => "show-legend",
        ChartConfigAction.SetStyle => "set-style",
        ChartConfigAction.SetPlacement => "set-placement",
        ChartConfigAction.SetDataLabels => "set-data-labels",
        ChartConfigAction.GetAxisScale => "get-axis-scale",
        ChartConfigAction.SetAxisScale => "set-axis-scale",
        ChartConfigAction.GetGridlines => "get-gridlines",
        ChartConfigAction.SetGridlines => "set-gridlines",
        ChartConfigAction.SetSeriesFormat => "set-series-format",
        ChartConfigAction.ListTrendlines => "list-trendlines",
        ChartConfigAction.AddTrendline => "add-trendline",
        ChartConfigAction.DeleteTrendline => "delete-trendline",
        ChartConfigAction.SetTrendline => "set-trendline",
        _ => throw new ArgumentException($"Unknown ChartConfigAction: {action}")
    };

    public static string ToActionString(this SlicerAction action) => action switch
    {
        SlicerAction.CreateSlicer => "create-slicer",
        SlicerAction.ListSlicers => "list-slicers",
        SlicerAction.SetSlicerSelection => "set-slicer-selection",
        SlicerAction.DeleteSlicer => "delete-slicer",
        SlicerAction.CreateTableSlicer => "create-table-slicer",
        SlicerAction.ListTableSlicers => "list-table-slicers",
        SlicerAction.SetTableSlicerSelection => "set-table-slicer-selection",
        SlicerAction.DeleteTableSlicer => "delete-table-slicer",
        _ => throw new ArgumentException($"Unknown SlicerAction: {action}")
    };
}
#pragma warning restore CS1591
