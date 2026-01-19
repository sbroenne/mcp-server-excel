namespace Sbroenne.ExcelMcp.McpServer.Models;

/// <summary>
/// Actions available for excel_file tool
/// </summary>
/// <remarks>
/// IMPORTANT: Keep enum values synchronized with ExcelFileTool.cs switch cases.
/// Enum names are PascalCase (CreateEmpty), converted to kebab-case (create-empty) via ActionExtensions.
/// Session Management: Open/Close manage persistent sessions. Close action has optional save parameter.
/// </remarks>
public enum FileAction
{
    List,
    Open,
    Close,
    CreateEmpty,
    CloseWorkbook,
    Test
}

/// <summary>
/// Actions available for excel_powerquery tool
/// </summary>
/// <remarks>
/// ATOMIC OPERATIONS: Improved workflow commands
/// - Create: Atomic import + load (replaces Import + SetLoadTo* + Refresh)
/// - UpdateMCode: Update formula only (explicit separation from refresh)
/// - LoadTo: Atomic configure + refresh (replaces SetLoadTo* + Refresh)
/// - Unload: Convert to connection-only (inverse of LoadTo)
/// - RefreshAll: Batch refresh all queries
///
/// NOTE: ValidateSyntax removed - Excel validation timing differs from test expectations
/// NOTE: UpdateMCode renamed to Update (auto-refreshes)
/// NOTE: UpdateAndRefresh removed (redundant - Update now auto-refreshes)
/// </remarks>
public enum PowerQueryAction
{
    List,
    View,
    Refresh,
    Delete,
    GetLoadConfig,

    // Atomic Operations
    Create,
    Update,       // Renamed from UpdateMCode, now auto-refreshes
    Rename,
    RefreshAll,
    LoadTo
}

/// <summary>
/// Actions available for excel_worksheet tool (lifecycle)
/// </summary>
/// <remarks>
/// ATOMIC CROSS-FILE OPERATIONS:
/// - CopyToFile: Copy sheet to another file (no session required)
/// - MoveToFile: Move sheet to another file (no session required)
/// These operations handle opening both files internally.
/// </remarks>
public enum WorksheetAction
{
    List,
    Create,
    Rename,
    Copy,
    Delete,
    Move,

    // Atomic cross-file operations
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
    // Values & Formulas
    GetValues,
    SetValues,
    GetFormulas,
    SetFormulas,

    // Number Formats
    GetNumberFormats,
    SetNumberFormat,
    SetNumberFormats,

    // Clear Operations
    ClearAll,
    ClearContents,
    ClearFormats,

    // Copy Operations
    Copy,
    CopyValues,
    CopyFormulas,

    // Discovery Operations
    GetUsedRange,
    GetCurrentRegion,
    GetInfo
}

/// <summary>
/// Actions available for excel_range_edit tool (insert/delete/search/sort)
/// </summary>
public enum RangeEditAction
{
    // Insert/Delete Cell Operations
    InsertCells,
    DeleteCells,

    // Insert/Delete Row Operations
    InsertRows,
    DeleteRows,

    // Insert/Delete Column Operations
    InsertColumns,
    DeleteColumns,

    // Search & Sort
    Find,
    Replace,
    Sort
}

/// <summary>
/// Actions available for excel_range_format tool (styling/validation/merge)
/// </summary>
public enum RangeFormatAction
{
    // Formatting & Validation
    GetStyle,
    SetStyle,
    FormatRange,
    ValidateRange,
    GetValidation,
    RemoveValidation,

    // Auto-Sizing
    AutoFitColumns,
    AutoFitRows,

    // Merge Operations
    MergeCells,
    UnmergeCells,
    GetMergeInfo
}

/// <summary>
/// Actions available for excel_range_link tool (hyperlinks/protection)
/// </summary>
public enum RangeLinkAction
{
    // Hyperlink Operations
    AddHyperlink,
    RemoveHyperlink,
    ListHyperlinks,
    GetHyperlink,

    // Cell Protection
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
    // Lifecycle
    List,
    Read,
    Create,
    Rename,
    Delete,
    Resize,

    // Styling & Totals
    SetStyle,
    ToggleTotals,
    SetColumnTotal,

    // Data Operations
    Append,
    GetData,

    // Data Model
    AddToDataModel,

    // DAX-Backed Tables
    CreateFromDax,
    UpdateDax,
    GetDax
}

/// <summary>
/// Actions available for excel_table_column tool (filter/column/sort)
/// </summary>
public enum TableColumnAction
{
    // Filter Operations
    ApplyFilter,
    ApplyFilterValues,
    ClearFilters,
    GetFilters,

    // Column Management
    AddColumn,
    RemoveColumn,
    RenameColumn,

    // Structured References
    GetStructuredReference,

    // Sort Operations
    Sort,
    SortMulti,

    // Number Formatting
    GetColumnNumberFormat,
    SetColumnNumberFormat
}

/// <summary>
/// Actions available for excel_pivottable tool (lifecycle operations)
/// </summary>
public enum PivotTableAction
{
    // Lifecycle
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
    // Field Management
    ListFields,
    AddRowField,
    AddColumnField,
    AddValueField,
    AddFilterField,
    RemoveField,

    // Field Configuration
    SetFieldFunction,
    SetFieldName,
    SetFieldFormat,
    SetFieldFilter,
    SortField,

    // Grouping Operations
    GroupByDate,
    GroupByNumeric
}

/// <summary>
/// Actions available for excel_pivottable_calc tool (calculated fields/members + layout + data)
/// </summary>
public enum PivotTableCalcAction
{
    // Calculated Fields (for regular PivotTables)
    ListCalculatedFields,
    CreateCalculatedField,
    DeleteCalculatedField,

    // Calculated Members (for OLAP/Data Model PivotTables)
    ListCalculatedMembers,
    CreateCalculatedMember,
    DeleteCalculatedMember,

    // Layout and Formatting
    SetLayout,
    SetSubtotals,
    SetGrandTotals,

    // Data Operations
    GetData
}

/// <summary>
/// Actions available for excel_chart tool (lifecycle)
/// </summary>
public enum ChartAction
{
    // Lifecycle
    List,
    Read,
    CreateFromRange,
    CreateFromPivotTable,
    Delete,
    Move
}

/// <summary>
/// Actions available for excel_chart_config tool (data source + appearance)
/// </summary>
public enum ChartConfigAction
{
    // Data Source Operations
    SetSourceRange,
    AddSeries,
    RemoveSeries,

    // Appearance Operations
    SetChartType,
    SetTitle,
    SetAxisTitle,
    ShowLegend,
    SetStyle
}

/// <summary>
/// Actions available for excel_slicer tool
/// </summary>
/// <remarks>
/// Slicers provide visual filtering controls for PivotTables and Tables.
/// A slicer can be connected to multiple PivotTables sharing the same SlicerCache.
/// Table slicers can only filter one Table.
/// </remarks>
public enum SlicerAction
{
    // === PivotTable Slicer Actions ===

    /// <summary>Create a slicer for a PivotTable field</summary>
    CreateSlicer,

    /// <summary>List all slicers in workbook, optionally filtered by PivotTable</summary>
    ListSlicers,

    /// <summary>Set or clear slicer selection (filter)</summary>
    SetSlicerSelection,

    /// <summary>Delete a slicer from the workbook</summary>
    DeleteSlicer,

    // === Table Slicer Actions ===

    /// <summary>Create a slicer for an Excel Table column</summary>
    CreateTableSlicer,

    /// <summary>List all Table slicers in workbook, optionally filtered by Table name</summary>
    ListTableSlicers,

    /// <summary>Set or clear Table slicer selection (filter)</summary>
    SetTableSlicerSelection,

    /// <summary>Delete a Table slicer from the workbook</summary>
    DeleteTableSlicer
}


