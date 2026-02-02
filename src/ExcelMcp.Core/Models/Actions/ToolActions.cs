#pragma warning disable CS1591
namespace Sbroenne.ExcelMcp.Core.Models.Actions;

/// <summary>
/// Actions available for excel_file tool
/// </summary>
/// <remarks>
/// IMPORTANT: Keep enum values synchronized with tool switch cases.
/// Enum names are PascalCase (e.g., Create), converted to kebab-case (e.g., create) via ActionExtensions.
///
/// </remarks>
public enum FileAction
{
    List,
    Open,
    Close,
    Create,
    CloseWorkbook,
    Test
}

/// <summary>
/// Actions available for excel_powerquery tool
/// </summary>
public enum PowerQueryAction
{
    List,
    View,
    Refresh,
    Delete,
    GetLoadConfig,
    Create,
    Update,
    Rename,
    RefreshAll,
    LoadTo,
    Unload,
    Evaluate
}

/// <summary>
/// Actions available for excel_worksheet tool (lifecycle)
/// </summary>
public enum WorksheetAction
{
    List,
    Create,
    Rename,
    Copy,
    Delete,
    Move,
    CopyToFile,
    MoveToFile
}

/// <summary>
/// Actions available for excel_worksheet_style tool (tab colors + visibility)
/// </summary>
public enum WorksheetStyleAction
{
    SetTabColor,
    GetTabColor,
    ClearTabColor,
    Hide,
    VeryHide,
    Show,
    GetVisibility,
    SetVisibility
}

/// <summary>
/// Actions available for excel_range tool (core data operations)
/// </summary>
public enum RangeAction
{
    GetValues,
    SetValues,
    GetFormulas,
    SetFormulas,
    GetNumberFormats,
    SetNumberFormat,
    SetNumberFormats,
    ClearAll,
    ClearContents,
    ClearFormats,
    Copy,
    CopyValues,
    CopyFormulas,
    GetUsedRange,
    GetCurrentRegion,
    GetInfo
}

/// <summary>
/// Actions available for excel_range_edit tool (insert/delete/search/sort)
/// </summary>
public enum RangeEditAction
{
    InsertCells,
    DeleteCells,
    InsertRows,
    DeleteRows,
    InsertColumns,
    DeleteColumns,
    Find,
    Replace,
    Sort
}

/// <summary>
/// Actions available for excel_range_format tool (styling/validation/merge)
/// </summary>
public enum RangeFormatAction
{
    GetStyle,
    SetStyle,
    FormatRange,
    ValidateRange,
    GetValidation,
    RemoveValidation,
    AutoFitColumns,
    AutoFitRows,
    MergeCells,
    UnmergeCells,
    GetMergeInfo
}

/// <summary>
/// Actions available for excel_range_link tool (hyperlinks/protection)
/// </summary>
public enum RangeLinkAction
{
    AddHyperlink,
    RemoveHyperlink,
    ListHyperlinks,
    GetHyperlink,
    SetCellLock,
    GetCellLock
}

/// <summary>
/// Actions available for excel_parameter tool
/// </summary>
public enum NamedRangeAction
{
    List,
    Read,
    Write,
    Create,
    Update,
    Delete
}

/// <summary>
/// Actions available for excel_conditional_format tool
/// </summary>
public enum ConditionalFormatAction
{
    AddRule,
    ClearRules
}

/// <summary>
/// Actions available for excel_vba tool
/// </summary>
public enum VbaAction
{
    List,
    View,
    Import,
    Delete,
    Run,
    Update
}

/// <summary>
/// Actions available for excel_connection tool
/// </summary>
public enum ConnectionAction
{
    List,
    View,
    Create,
    Test,
    Refresh,
    Delete,
    LoadTo,
    GetProperties,
    SetProperties
}

/// <summary>
/// Actions available for excel_datamodel tool (tables + measures)
/// </summary>
public enum DataModelAction
{
    ListTables,
    ReadTable,
    ListColumns,
    ListMeasures,
    Read,
    CreateMeasure,
    UpdateMeasure,
    DeleteMeasure,
    RenameTable,
    DeleteTable,
    ReadInfo,
    Refresh,
    Evaluate,
    ExecuteDmv
}

/// <summary>
/// Actions available for excel_datamodel_rel tool (relationships)
/// </summary>
public enum DataModelRelAction
{
    ListRelationships,
    ReadRelationship,
    CreateRelationship,
    UpdateRelationship,
    DeleteRelationship
}

/// <summary>
/// Actions available for excel_table tool (core lifecycle/data)
/// </summary>
public enum TableAction
{
    List,
    Read,
    Create,
    Rename,
    Delete,
    Resize,
    SetStyle,
    ToggleTotals,
    SetColumnTotal,
    Append,
    GetData,
    AddToDataModel,
    CreateFromDax,
    UpdateDax,
    GetDax
}

/// <summary>
/// Actions available for excel_table_column tool (filter/column/sort)
/// </summary>
public enum TableColumnAction
{
    ApplyFilter,
    ApplyFilterValues,
    ClearFilters,
    GetFilters,
    AddColumn,
    RemoveColumn,
    RenameColumn,
    GetStructuredReference,
    Sort,
    SortMulti,
    GetColumnNumberFormat,
    SetColumnNumberFormat
}

/// <summary>
/// Actions available for excel_pivottable tool (lifecycle operations)
/// </summary>
public enum PivotTableAction
{
    List,
    Read,
    CreateFromRange,
    CreateFromTable,
    CreateFromDataModel,
    Delete,
    Refresh
}

/// <summary>
/// Actions available for excel_pivottable_field tool (field management)
/// </summary>
public enum PivotTableFieldAction
{
    ListFields,
    AddRowField,
    AddColumnField,
    AddValueField,
    AddFilterField,
    RemoveField,
    SetFieldFunction,
    SetFieldName,
    SetFieldFormat,
    SetFieldFilter,
    SortField,
    GroupByDate,
    GroupByNumeric
}

/// <summary>
/// Actions available for excel_pivottable_calc tool (calculated fields/members + layout + data)
/// </summary>
public enum PivotTableCalcAction
{
    ListCalculatedFields,
    CreateCalculatedField,
    DeleteCalculatedField,
    ListCalculatedMembers,
    CreateCalculatedMember,
    DeleteCalculatedMember,
    SetLayout,
    SetSubtotals,
    SetGrandTotals,
    GetData
}

/// <summary>
/// Actions available for excel_chart tool (lifecycle)
/// </summary>
public enum ChartAction
{
    List,
    Read,
    CreateFromRange,
    CreateFromTable,
    CreateFromPivotTable,
    Delete,
    Move,
    FitToRange
}

/// <summary>
/// Actions available for excel_chart_config tool (data source + appearance)
/// </summary>
public enum ChartConfigAction
{
    SetSourceRange,
    AddSeries,
    RemoveSeries,
    SetChartType,
    SetTitle,
    SetAxisTitle,
    GetAxisNumberFormat,
    SetAxisNumberFormat,
    ShowLegend,
    SetStyle,
    SetPlacement,
    SetDataLabels,
    GetAxisScale,
    SetAxisScale,
    GetGridlines,
    SetGridlines,
    SetSeriesFormat,
    ListTrendlines,
    AddTrendline,
    DeleteTrendline,
    SetTrendline
}

/// <summary>
/// Actions available for excel_slicer tool
/// </summary>
public enum SlicerAction
{
    CreateSlicer,
    ListSlicers,
    SetSlicerSelection,
    DeleteSlicer,
    CreateTableSlicer,
    ListTableSlicers,
    SetTableSlicerSelection,
    DeleteTableSlicer
}
#pragma warning restore CS1591
