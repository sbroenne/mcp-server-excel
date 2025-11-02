namespace Sbroenne.ExcelMcp.McpServer.Models;

/// <summary>
/// Actions available for excel_file tool
/// </summary>
/// <remarks>
/// IMPORTANT: Keep enum values synchronized with ExcelFileTool.cs switch cases.
/// Enum names are PascalCase (CreateEmpty), converted to kebab-case (create-empty) via ActionExtensions.
/// </remarks>
public enum FileAction
{
    CreateEmpty,
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
    Import,
    Export,
    Update,
    Refresh,
    Delete,
    SetLoadToTable,
    SetLoadToDataModel,
    SetLoadToBoth,
    SetConnectionOnly,
    GetLoadConfig
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
    GetRangeInfo,
    
    // Hyperlink Operations
    AddHyperlink,
    RemoveHyperlink,
    ListHyperlinks,
    GetHyperlink,
    
    // Formatting & Validation
    FormatRange,
    ValidateRange
}

/// <summary>
/// Actions available for excel_parameter tool
/// </summary>
public enum ParameterAction
{
    List,
    Create,
    CreateBulk,
    Update,
    Delete,
    Get,
    Set
}

/// <summary>
/// Actions available for excel_vba tool
/// </summary>
public enum VbaAction
{
    List,
    View,
    Import,
    Export,
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
    Import,
    Export,
    UpdateProperties,
    Test,
    Refresh,
    Delete
}

/// <summary>
/// Actions available for excel_datamodel tool
/// </summary>
public enum DataModelAction
{
    ListTables,
    ViewTable,
    ListColumns,
    ListMeasures,
    ViewMeasure,
    CreateMeasure,
    UpdateMeasure,
    DeleteMeasure,
    ListRelationships,
    CreateRelationship,
    DeleteRelationship,
    GetModelInfo,
    Refresh
}

/// <summary>
/// Actions available for excel_table tool
/// </summary>
public enum TableAction
{
    List,
    Info,
    Create,
    Rename,
    Delete,
    Resize,
    GetStructuredReference,
    AddToDataModel
}

/// <summary>
/// Actions available for excel_pivottable tool
/// </summary>
public enum PivotTableAction
{
    CreateFromRange,
    CreateFromTable,
    ListFields,
    AddRowField,
    AddColumnField,
    AddDataField,
    AddFilterField,
    Refresh
}

