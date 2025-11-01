namespace Sbroenne.ExcelMcp.McpServer.Models;

/// <summary>
/// Helper extensions to convert enum actions to string format expected by Core commands
/// </summary>
public static class ActionExtensions
{
    public static string ToActionString(this FileAction action) => action switch
    {
        FileAction.CreateEmpty => "create-empty",
        FileAction.CloseWorkbook => "close-workbook",
        FileAction.Test => "test",
        _ => throw new ArgumentException($"Unknown FileAction: {action}")
    };

    public static string ToActionString(this PowerQueryAction action) => action switch
    {
        PowerQueryAction.List => "list",
        PowerQueryAction.View => "view",
        PowerQueryAction.Import => "import",
        PowerQueryAction.Export => "export",
        PowerQueryAction.Update => "update",
        PowerQueryAction.Refresh => "refresh",
        PowerQueryAction.Delete => "delete",
        PowerQueryAction.SetLoadToTable => "set-load-to-table",
        PowerQueryAction.SetLoadToDataModel => "set-load-to-data-model",
        PowerQueryAction.SetLoadToBoth => "set-load-to-both",
        PowerQueryAction.SetConnectionOnly => "set-connection-only",
        PowerQueryAction.GetLoadConfig => "get-load-config",
        _ => throw new ArgumentException($"Unknown PowerQueryAction: {action}")
    };

    public static string ToActionString(this WorksheetAction action) => action switch
    {
        WorksheetAction.List => "list",
        WorksheetAction.Create => "create",
        WorksheetAction.Rename => "rename",
        WorksheetAction.Copy => "copy",
        WorksheetAction.Delete => "delete",
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
        RangeAction.ClearContents => "clear-contents",
        RangeAction.ClearAll => "clear-all",
        RangeAction.GetInfo => "get-info",
        RangeAction.Copy => "copy",
        RangeAction.CopyValues => "copy-values",
        RangeAction.Find => "find",
        RangeAction.Replace => "replace",
        RangeAction.Sort => "sort",
        RangeAction.GetNumberFormats => "get-number-formats",
        RangeAction.SetNumberFormat => "set-number-format",
        RangeAction.SetNumberFormats => "set-number-formats",
        RangeAction.AddHyperlink => "add-hyperlink",
        RangeAction.RemoveHyperlink => "remove-hyperlink",
        RangeAction.GetHyperlinks => "get-hyperlinks",
        _ => throw new ArgumentException($"Unknown RangeAction: {action}")
    };

    public static string ToActionString(this ParameterAction action) => action switch
    {
        ParameterAction.List => "list",
        ParameterAction.Create => "create",
        ParameterAction.Delete => "delete",
        ParameterAction.Get => "get",
        ParameterAction.Set => "set",
        _ => throw new ArgumentException($"Unknown ParameterAction: {action}")
    };

    public static string ToActionString(this VbaAction action) => action switch
    {
        VbaAction.List => "list",
        VbaAction.View => "view",
        VbaAction.Import => "import",
        VbaAction.Export => "export",
        VbaAction.Delete => "delete",
        VbaAction.Run => "run",
        VbaAction.Update => "update",
        _ => throw new ArgumentException($"Unknown VbaAction: {action}")
    };

    public static string ToActionString(this ConnectionAction action) => action switch
    {
        ConnectionAction.List => "list",
        ConnectionAction.View => "view",
        ConnectionAction.Import => "import",
        ConnectionAction.Export => "export",
        ConnectionAction.UpdateProperties => "update-properties",
        ConnectionAction.Test => "test",
        ConnectionAction.Refresh => "refresh",
        ConnectionAction.Delete => "delete",
        _ => throw new ArgumentException($"Unknown ConnectionAction: {action}")
    };

    public static string ToActionString(this DataModelAction action) => action switch
    {
        DataModelAction.ListTables => "list-tables",
        DataModelAction.ViewTable => "view-table",
        DataModelAction.ListColumns => "list-columns",
        DataModelAction.ListMeasures => "list-measures",
        DataModelAction.ViewMeasure => "view-measure",
        DataModelAction.CreateMeasure => "create-measure",
        DataModelAction.UpdateMeasure => "update-measure",
        DataModelAction.DeleteMeasure => "delete-measure",
        DataModelAction.ListRelationships => "list-relationships",
        DataModelAction.CreateRelationship => "create-relationship",
        DataModelAction.DeleteRelationship => "delete-relationship",
        DataModelAction.GetModelInfo => "get-model-info",
        DataModelAction.Refresh => "refresh",
        _ => throw new ArgumentException($"Unknown DataModelAction: {action}")
    };

    public static string ToActionString(this TableAction action) => action switch
    {
        TableAction.List => "list",
        TableAction.Info => "info",
        TableAction.Create => "create",
        TableAction.Rename => "rename",
        TableAction.Delete => "delete",
        TableAction.Resize => "resize",
        TableAction.GetStructuredReference => "get-structured-reference",
        TableAction.AddToDataModel => "add-to-datamodel",
        _ => throw new ArgumentException($"Unknown TableAction: {action}")
    };

    public static string ToActionString(this PivotTableAction action) => action switch
    {
        PivotTableAction.CreateFromRange => "create-from-range",
        PivotTableAction.CreateFromTable => "create-from-table",
        PivotTableAction.ListFields => "list-fields",
        PivotTableAction.AddRowField => "add-row-field",
        PivotTableAction.AddColumnField => "add-column-field",
        PivotTableAction.AddDataField => "add-data-field",
        PivotTableAction.AddFilterField => "add-filter-field",
        PivotTableAction.Refresh => "refresh",
        _ => throw new ArgumentException($"Unknown PivotTableAction: {action}")
    };
}
