namespace Sbroenne.ExcelMcp.McpServer.Models;

/// <summary>
/// Helper extensions to convert enum actions to string format expected by Core commands
/// </summary>
public static class ActionExtensions
{
    public static string ToActionString(this FileAction action) => action switch
    {
        FileAction.Open => "open",
        FileAction.Save => "save",
        FileAction.Close => "close",
        FileAction.CreateEmpty => "create-empty",
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
        PowerQueryAction.LoadTo => "load-to",

        // Atomic Operations
        PowerQueryAction.Create => "create",
        PowerQueryAction.Update => "update",  // Renamed from update-mcode
        PowerQueryAction.RefreshAll => "refresh-all",

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
        WorksheetAction.CopyToWorkbook => "copy-to-workbook",
        WorksheetAction.MoveToWorkbook => "move-to-workbook",
        WorksheetAction.SetTabColor => "set-tab-color",
        WorksheetAction.GetTabColor => "get-tab-color",
        WorksheetAction.ClearTabColor => "clear-tab-color",
        WorksheetAction.Hide => "hide",
        WorksheetAction.VeryHide => "very-hide",
        WorksheetAction.Show => "show",
        WorksheetAction.GetVisibility => "get-visibility",
        WorksheetAction.SetVisibility => "set-visibility",
        _ => throw new ArgumentException($"Unknown WorksheetAction: {action}")
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
        RangeAction.InsertCells => "insert-cells",
        RangeAction.DeleteCells => "delete-cells",
        RangeAction.InsertRows => "insert-rows",
        RangeAction.DeleteRows => "delete-rows",
        RangeAction.InsertColumns => "insert-columns",
        RangeAction.DeleteColumns => "delete-columns",
        RangeAction.Find => "find",
        RangeAction.Replace => "replace",
        RangeAction.Sort => "sort",
        RangeAction.GetUsedRange => "get-used-range",
        RangeAction.GetCurrentRegion => "get-current-region",
        RangeAction.GetInfo => "get-info",
        RangeAction.AddHyperlink => "add-hyperlink",
        RangeAction.RemoveHyperlink => "remove-hyperlink",
        RangeAction.ListHyperlinks => "list-hyperlinks",
        RangeAction.GetHyperlink => "get-hyperlink",
        RangeAction.GetStyle => "get-style",
        RangeAction.SetStyle => "set-style",
        RangeAction.FormatRange => "format-range",
        RangeAction.ValidateRange => "validate-range",
        RangeAction.GetValidation => "get-validation",
        RangeAction.RemoveValidation => "remove-validation",
        RangeAction.AutoFitColumns => "auto-fit-columns",
        RangeAction.AutoFitRows => "auto-fit-rows",
        RangeAction.MergeCells => "merge-cells",
        RangeAction.UnmergeCells => "unmerge-cells",
        RangeAction.GetMergeInfo => "get-merge-info",
        RangeAction.SetCellLock => "set-cell-lock",
        RangeAction.GetCellLock => "get-cell-lock",
        _ => throw new ArgumentException($"Unknown RangeAction: {action}")
    };

    public static string ToActionString(this NamedRangeAction action) => action switch
    {
        NamedRangeAction.List => "list",
        NamedRangeAction.Read => "read",
        NamedRangeAction.Write => "write",
        NamedRangeAction.Create => "create",
        NamedRangeAction.CreateBulk => "create-bulk",
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
        DataModelAction.ListRelationships => "list-relationships",
        DataModelAction.CreateRelationship => "create-relationship",
        DataModelAction.UpdateRelationship => "update-relationship",
        DataModelAction.DeleteRelationship => "delete-relationship",
        DataModelAction.ReadInfo => "read-info",
        DataModelAction.Refresh => "refresh",
        _ => throw new ArgumentException($"Unknown DataModelAction: {action}")
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
        TableAction.AddToDataModel => "add-to-datamodel",
        TableAction.ApplyFilter => "apply-filter",
        TableAction.ApplyFilterValues => "apply-filter-values",
        TableAction.ClearFilters => "clear-filters",
        TableAction.GetFilters => "get-filters",
        TableAction.AddColumn => "add-column",
        TableAction.RemoveColumn => "remove-column",
        TableAction.RenameColumn => "rename-column",
        TableAction.GetStructuredReference => "get-structured-reference",
        TableAction.Sort => "sort",
        TableAction.SortMulti => "sort-multi",
        TableAction.GetColumnNumberFormat => "get-column-number-format",
        TableAction.SetColumnNumberFormat => "set-column-number-format",
        _ => throw new ArgumentException($"Unknown TableAction: {action}")
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
        PivotTableAction.ListFields => "list-fields",
        PivotTableAction.AddRowField => "add-row-field",
        PivotTableAction.AddColumnField => "add-column-field",
        PivotTableAction.AddValueField => "add-value-field",
        PivotTableAction.AddFilterField => "add-filter-field",
        PivotTableAction.RemoveField => "remove-field",
        PivotTableAction.SetFieldFunction => "set-field-function",
        PivotTableAction.SetFieldName => "set-field-name",
        PivotTableAction.SetFieldFormat => "set-field-format",
        PivotTableAction.SetFieldFilter => "set-field-filter",
        PivotTableAction.SortField => "sort-field",
        PivotTableAction.GetData => "get-data",
        _ => throw new ArgumentException($"Unknown PivotTableAction: {action}")
    };

    public static string ToActionString(this QueryTableAction action) => action switch
    {
        QueryTableAction.List => "list",
        QueryTableAction.Read => "read",
        QueryTableAction.CreateFromConnection => "create-from-connection",
        QueryTableAction.CreateFromQuery => "create-from-query",
        QueryTableAction.Refresh => "refresh",
        QueryTableAction.RefreshAll => "refresh-all",
        QueryTableAction.UpdateProperties => "update-properties",
        QueryTableAction.Delete => "delete",
        _ => throw new ArgumentException($"Unknown QueryTableAction: {action}")
    };
}


