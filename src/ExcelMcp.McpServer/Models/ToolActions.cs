namespace Sbroenne.ExcelMcp.McpServer.Models;

/// <summary>
/// Actions available for excel_file tool
/// </summary>
/// <remarks>
/// IMPORTANT: Keep enum values synchronized with ExcelFileTool.cs switch cases.
/// Enum names are PascalCase (CreateEmpty), converted to kebab-case (create-empty) via ActionExtensions.
/// Session Management: Open/Save/Close manage persistent sessions across multiple operations.
/// </remarks>
public enum FileAction
{
    Open,
    Save,
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
    ListExcelSources,

    // Atomic Operations
    Create,
    Update,       // Renamed from UpdateMCode, now auto-refreshes
    Unload,
    RefreshAll,
    LoadTo
}

/// <summary>
/// Actions available for excel_worksheet tool
/// </summary>
public enum WorksheetAction
{
    List,
    Create,
    Rename,
    Copy,
    Delete,
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
/// Actions available for excel_range tool
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
    Sort,

    // Discovery Operations
    GetUsedRange,
    GetCurrentRegion,
    GetInfo,

    // Hyperlink Operations
    AddHyperlink,
    RemoveHyperlink,
    ListHyperlinks,
    GetHyperlink,

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
    GetMergeInfo,

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
    CreateBulk,
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
    Import,
    UpdateProperties,
    Test,
    Refresh,
    Delete,
    LoadTo,
    GetProperties,
    SetProperties
}

/// <summary>
/// Actions available for excel_datamodel tool
/// </summary>
public enum DataModelAction
{
    ListTables,
    ReadTable,
    ListColumns,
    ListMeasures,
    Read,
    ExportMeasure,
    CreateMeasure,
    UpdateMeasure,
    DeleteMeasure,
    ListRelationships,
    CreateRelationship,
    UpdateRelationship,
    DeleteRelationship,
    ReadInfo,
    Refresh
}

/// <summary>
/// Actions available for excel_table tool
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

    // Data Model
    AddToDataModel,

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
/// Actions available for excel_pivottable tool
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
    Refresh,

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

    // Data Operations
    GetData
}

/// <summary>
/// Actions available for excel_querytable tool
/// </summary>
public enum QueryTableAction
{
    List,
    Read,
    CreateFromConnection,
    CreateFromQuery,
    Refresh,
    RefreshAll,
    UpdateProperties,
    Delete
}


