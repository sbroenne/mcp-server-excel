using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Base result type for all Core operations.
/// NOTE: Core commands should NOT set SuggestedNextActions (workflow guidance is MCP/CLI layer responsibility).
/// Exceptions propagate naturally to batch.Execute() which converts them to OperationResult { Success = false }.
/// </summary>
/// <remarks>
/// Property names are intentionally short to minimize JSON token count for LLM efficiency:
/// - ok: Success
/// - err: ErrorMessage
/// - fp: FilePath
/// </remarks>
public abstract class ResultBase
{
    /// <summary>
    /// Indicates whether the operation was successful
    /// </summary>
    [JsonPropertyName("ok")]
    public bool Success { get; set; }

    /// <summary>
    /// Error message if operation failed
    /// </summary>
    [JsonPropertyName("err")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ErrorMessage { get; set; }

    /// <summary>
    /// File path of the Excel file
    /// </summary>
    [JsonPropertyName("fp")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FilePath { get; set; }
}

/// <summary>
/// Result for operations that don't return data (create, delete, etc.)
/// </summary>
public class OperationResult : ResultBase
{
    /// <summary>
    /// Action that was performed
    /// </summary>
    [JsonPropertyName("act")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Action { get; set; }
}

/// <summary>
/// Result for rename operations across Core features
/// </summary>
/// <remarks>
/// Property names: ot=ObjectType, on=OldName, nn=NewName, non=NormalizedOldName, nnn=NormalizedNewName
/// </remarks>
public class RenameResult : ResultBase
{
    /// <summary>
    /// Type of object being renamed (power-query, data-model-table)
    /// </summary>
    [JsonPropertyName("ot")]
    public string ObjectType { get; set; } = string.Empty;

    /// <summary>
    /// Original name provided by the caller
    /// </summary>
    [JsonPropertyName("on")]
    public string OldName { get; set; } = string.Empty;

    /// <summary>
    /// Desired new name provided by the caller
    /// </summary>
    [JsonPropertyName("nn")]
    public string NewName { get; set; } = string.Empty;

    /// <summary>
    /// Trimmed old name used for comparisons
    /// </summary>
    [JsonPropertyName("non")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string NormalizedOldName { get; set; } = string.Empty;

    /// <summary>
    /// Trimmed new name used for comparisons
    /// </summary>
    [JsonPropertyName("nnn")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string NormalizedNewName { get; set; } = string.Empty;
}

/// <summary>
/// Result for listing worksheets
/// </summary>
/// <remarks>
/// Property names: ws=Worksheets
/// </remarks>
public class WorksheetListResult : ResultBase
{
    /// <summary>
    /// List of worksheets in the workbook
    /// </summary>
    [JsonPropertyName("ws")]
    public List<WorksheetInfo> Worksheets { get; set; } = [];
}

/// <summary>
/// Information about a worksheet
/// </summary>
/// <remarks>
/// Property names: n=Name, i=Index, v=Visible
/// </remarks>
public class WorksheetInfo
{
    /// <summary>
    /// Name of the worksheet
    /// </summary>
    [JsonPropertyName("n")]
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Index of the worksheet (1-based)
    /// </summary>
    [JsonPropertyName("i")]
    public int Index { get; set; }

    /// <summary>
    /// Whether the worksheet is visible
    /// </summary>
    [JsonPropertyName("v")]
    public bool Visible { get; set; }
}

/// <summary>
/// Sheet visibility levels (maps to Excel XlSheetVisibility)
/// </summary>
public enum SheetVisibility
{
    /// <summary>
    /// Sheet is visible (xlSheetVisible = -1)
    /// </summary>
    Visible = -1,

    /// <summary>
    /// Sheet is hidden but user can unhide via Excel UI (xlSheetHidden = 0)
    /// </summary>
    Hidden = 0,

    /// <summary>
    /// Sheet is very hidden, requires code to unhide (xlSheetVeryHidden = 2)
    /// </summary>
    VeryHidden = 2
}

/// <summary>
/// Result for getting worksheet tab color
/// </summary>
public class TabColorResult : ResultBase
{
    /// <summary>
    /// Whether the sheet has a tab color set
    /// </summary>
    public bool HasColor { get; set; }

    /// <summary>
    /// Red component (0-255), null if no color
    /// </summary>
    public int? Red { get; set; }

    /// <summary>
    /// Green component (0-255), null if no color
    /// </summary>
    public int? Green { get; set; }

    /// <summary>
    /// Blue component (0-255), null if no color
    /// </summary>
    public int? Blue { get; set; }

    /// <summary>
    /// Hex color string (#RRGGBB), null if no color
    /// </summary>
    public string? HexColor { get; set; }
}

/// <summary>
/// Result for getting worksheet visibility
/// </summary>
public class SheetVisibilityResult : ResultBase
{
    /// <summary>
    /// Visibility level
    /// </summary>
    public SheetVisibility Visibility { get; set; }

    /// <summary>
    /// Visibility name (Visible, Hidden, VeryHidden)
    /// </summary>
    public string VisibilityName { get; set; } = string.Empty;
}

/// <summary>
/// Result for reading worksheet data
/// </summary>
public class WorksheetDataResult : ResultBase
{
    /// <summary>
    /// Name of the worksheet
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Range that was read
    /// </summary>
    public string Range { get; set; } = string.Empty;

    /// <summary>
    /// Data rows and columns
    /// </summary>
    public List<List<object?>> Data { get; set; } = [];

    /// <summary>
    /// Column headers
    /// </summary>
    public List<string> Headers { get; set; } = [];

    /// <summary>
    /// Number of rows
    /// </summary>
    public int RowCount { get; set; }

    /// <summary>
    /// Number of columns
    /// </summary>
    public int ColumnCount { get; set; }
}

/// <summary>
/// Result for listing Power Queries
/// </summary>
public class PowerQueryListResult : ResultBase
{
    /// <summary>
    /// List of Power Queries in the workbook
    /// </summary>
    public List<PowerQueryInfo> Queries { get; set; } = [];
}

/// <summary>
/// Information about a Power Query
/// </summary>
public class PowerQueryInfo
{
    /// <summary>
    /// Name of the Power Query
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Full M code formula
    /// </summary>
    public string Formula { get; set; } = string.Empty;

    /// <summary>
    /// Preview of the formula (first 80 characters)
    /// </summary>
    public string FormulaPreview { get; set; } = string.Empty;

    /// <summary>
    /// Whether the query is connection-only
    /// </summary>
    public bool IsConnectionOnly { get; set; }
}

/// <summary>
/// Result for viewing Power Query code
/// </summary>
public class PowerQueryViewResult : ResultBase
{
    /// <summary>
    /// Name of the Power Query
    /// </summary>
    public string QueryName { get; set; } = string.Empty;

    /// <summary>
    /// Full M code
    /// </summary>
    public string MCode { get; set; } = string.Empty;

    /// <summary>
    /// Number of characters in the M code
    /// </summary>
    public int CharacterCount { get; set; }

    /// <summary>
    /// Whether the query is connection-only
    /// </summary>
    public bool IsConnectionOnly { get; set; }
}

/// <summary>
/// Power Query load configuration modes
/// </summary>
public enum PowerQueryLoadMode
{
    /// <summary>
    /// Connection only - no data loaded to worksheet or data model
    /// </summary>
    ConnectionOnly,

    /// <summary>
    /// Load to table in worksheet
    /// </summary>
    LoadToTable,

    /// <summary>
    /// Load to Data Model (PowerPivot)
    /// </summary>
    LoadToDataModel,

    /// <summary>
    /// Load to both table and data model
    /// </summary>
    LoadToBoth
}

/// <summary>
/// Result for Power Query load configuration
/// </summary>
public class PowerQueryLoadConfigResult : ResultBase
{
    /// <summary>
    /// Name of the query
    /// </summary>
    public string QueryName { get; set; } = string.Empty;

    /// <summary>
    /// Current load mode
    /// </summary>
    public PowerQueryLoadMode LoadMode { get; set; }

    /// <summary>
    /// Target worksheet name (if LoadToTable or LoadToBoth)
    /// </summary>
    public string? TargetSheet { get; set; }

    /// <summary>
    /// Whether the query has an active connection
    /// </summary>
    public bool HasConnection { get; set; }

    /// <summary>
    /// Whether the query is loaded to data model
    /// </summary>
    public bool IsLoadedToDataModel { get; set; }
}

/// <summary>
/// Result for Power Query load-to-data-model operations with verification
/// Extends OperationResult to provide detailed verification of atomic operation
/// </summary>
public class PowerQueryLoadToDataModelResult : OperationResult
{
    /// <summary>
    /// Name of the query
    /// </summary>
    public string QueryName { get; set; } = string.Empty;

    /// <summary>
    /// Whether the load configuration was successfully applied
    /// </summary>
    public bool ConfigurationApplied { get; set; }

    /// <summary>
    /// Whether data was actually loaded to the Data Model
    /// </summary>
    public bool DataLoadedToModel { get; set; }

    /// <summary>
    /// Number of rows loaded to the Data Model (0 if not loaded)
    /// </summary>
    public int RowsLoaded { get; set; }

    /// <summary>
    /// Total number of tables in the Data Model after operation
    /// </summary>
    public int TablesInDataModel { get; set; }

    /// <summary>
    /// Overall workflow status: "Complete" | "Failed" | "Partial"
    /// </summary>
    public string WorkflowStatus { get; set; } = "Failed";
}

/// <summary>
/// Result for Power Query load-to-table operations with verification
/// Extends OperationResult to provide detailed verification of atomic operation
/// </summary>
public class PowerQueryLoadToTableResult : OperationResult
{
    /// <summary>
    /// Name of the query
    /// </summary>
    public string QueryName { get; set; } = string.Empty;

    /// <summary>
    /// Name of the target worksheet
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Whether the load configuration was successfully applied
    /// </summary>
    public bool ConfigurationApplied { get; set; }

    /// <summary>
    /// Whether data was actually loaded to the worksheet table
    /// </summary>
    public bool DataLoadedToTable { get; set; }

    /// <summary>
    /// Number of rows loaded to the worksheet table (0 if not loaded)
    /// </summary>
    public int RowsLoaded { get; set; }

    /// <summary>
    /// Overall workflow status: "Complete" | "Failed" | "Partial"
    /// </summary>
    public string WorkflowStatus { get; set; } = "Failed";
}

/// <summary>
/// Result for Power Query load-to-both operations with verification
/// Extends OperationResult to provide detailed verification of atomic operation
/// </summary>
public class PowerQueryLoadToBothResult : OperationResult
{
    /// <summary>
    /// Name of the query
    /// </summary>
    public string QueryName { get; set; } = string.Empty;

    /// <summary>
    /// Name of the target worksheet
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Whether the load configuration was successfully applied
    /// </summary>
    public bool ConfigurationApplied { get; set; }

    /// <summary>
    /// Whether data was actually loaded to the worksheet table
    /// </summary>
    public bool DataLoadedToTable { get; set; }

    /// <summary>
    /// Whether data was actually loaded to the Data Model
    /// </summary>
    public bool DataLoadedToModel { get; set; }

    /// <summary>
    /// Number of rows loaded to the worksheet table (0 if not loaded)
    /// </summary>
    public int RowsLoadedToTable { get; set; }

    /// <summary>
    /// Number of rows loaded to the Data Model (0 if not loaded)
    /// </summary>
    public int RowsLoadedToModel { get; set; }

    /// <summary>
    /// Total number of tables in the Data Model after operation
    /// </summary>
    public int TablesInDataModel { get; set; }

    /// <summary>
    /// Overall workflow status: "Complete" | "Failed" | "Partial"
    /// </summary>
    public string WorkflowStatus { get; set; } = "Failed";
}

/// <summary>
/// Result for Power Query create operations
/// Atomic operation: Import M code + Load data to destination in ONE call
/// </summary>
public class PowerQueryCreateResult : OperationResult
{
    /// <summary>
    /// Name of the created query
    /// </summary>
    public string QueryName { get; set; } = string.Empty;

    /// <summary>
    /// Load destination applied
    /// </summary>
    public PowerQueryLoadMode LoadDestination { get; set; }

    /// <summary>
    /// Target worksheet name (if LoadToTable or LoadToBoth)
    /// </summary>
    public string? WorksheetName { get; set; }

    /// <summary>
    /// Target cell address used when loading to a worksheet (e.g., "A1")
    /// </summary>
    public string? TargetCellAddress { get; set; }

    /// <summary>
    /// Whether the query was created successfully
    /// </summary>
    public bool QueryCreated { get; set; }

    /// <summary>
    /// Whether data was loaded (true for all except ConnectionOnly)
    /// </summary>
    public bool DataLoaded { get; set; }

    /// <summary>
    /// Number of rows loaded (0 if ConnectionOnly)
    /// </summary>
    public int RowsLoaded { get; set; }
}

/// <summary>
/// Result for Power Query load operations
/// Atomic operation: Set destination + Refresh data in ONE call
/// </summary>
public class PowerQueryLoadResult : OperationResult
{
    /// <summary>
    /// Name of the query
    /// </summary>
    public string QueryName { get; set; } = string.Empty;

    /// <summary>
    /// Load destination applied
    /// </summary>
    public PowerQueryLoadMode LoadDestination { get; set; }

    /// <summary>
    /// Target worksheet name (if applicable)
    /// </summary>
    public string? WorksheetName { get; set; }

    /// <summary>
    /// Target cell address used for the worksheet load destination (null defaults to A1)
    /// </summary>
    public string? TargetCellAddress { get; set; }

    /// <summary>
    /// Whether load configuration was applied
    /// </summary>
    public bool ConfigurationApplied { get; set; }

    /// <summary>
    /// Whether data was refreshed
    /// </summary>
    public bool DataRefreshed { get; set; }

    /// <summary>
    /// Number of rows loaded
    /// </summary>
    public int RowsLoaded { get; set; }
}

/// <summary>
/// Result for Power Query syntax validation
/// Pre-flight syntax check before creating permanent query
/// </summary>
public class PowerQueryValidationResult : ResultBase
{
    /// <summary>
    /// Whether the M code syntax is valid
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// Validation errors (if any)
    /// </summary>
    public List<string> ValidationErrors { get; set; } = [];

    /// <summary>
    /// M code expression that was validated
    /// </summary>
    public string? MCodeExpression { get; set; }
}

/// <summary>
/// Result for listing named ranges/parameters
/// </summary>
public class NamedRangeListResult : ResultBase
{
    /// <summary>
    /// List of named ranges/parameters
    /// </summary>
    public List<NamedRangeInfo> NamedRanges { get; set; } = [];
}

/// <summary>
/// Information about a named range/parameter
/// </summary>
public class NamedRangeInfo
{
    /// <summary>
    /// Name of the named range
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// What the named range refers to
    /// </summary>
    public string RefersTo { get; set; } = string.Empty;

    /// <summary>
    /// Current value
    /// </summary>
    public object? Value { get; set; }

    /// <summary>
    /// Type of the value
    /// </summary>
    public string ValueType { get; set; } = string.Empty;
}

/// <summary>
/// Named range value information (for Read operation)
/// </summary>
public class NamedRangeValue
{
    /// <summary>
    /// Name of the named range
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// What the named range refers to
    /// </summary>
    public string RefersTo { get; set; } = string.Empty;

    /// <summary>
    /// Current value
    /// </summary>
    public object? Value { get; set; }

    /// <summary>
    /// Type of the value
    /// </summary>
    public string ValueType { get; set; } = string.Empty;
}

/// <summary>
/// Result for getting parameter value
/// </summary>
public class NamedRangeValueResult : ResultBase
{
    /// <summary>
    /// Name of the named range
    /// </summary>
    public string NamedRangeName { get; set; } = string.Empty;

    /// <summary>
    /// Current value
    /// </summary>
    public object? Value { get; set; }

    /// <summary>
    /// Type of the value
    /// </summary>
    public string ValueType { get; set; } = string.Empty;

    /// <summary>
    /// What the named range refers to
    /// </summary>
    public string RefersTo { get; set; } = string.Empty;
}

/// <summary>
/// Result for listing VBA scripts
/// </summary>
public class VbaListResult : ResultBase
{
    /// <summary>
    /// List of VBA scripts
    /// </summary>
    public List<ScriptInfo> Scripts { get; set; } = [];
}

/// <summary>
/// Result for viewing VBA module code
/// </summary>
public class VbaViewResult : ResultBase
{
    /// <summary>
    /// Module name
    /// </summary>
    public string ModuleName { get; set; } = string.Empty;

    /// <summary>
    /// Module type
    /// </summary>
    public string ModuleType { get; set; } = string.Empty;

    /// <summary>
    /// Complete VBA code
    /// </summary>
    public string Code { get; set; } = string.Empty;

    /// <summary>
    /// Number of lines in the module
    /// </summary>
    public int LineCount { get; set; }

    /// <summary>
    /// List of procedures in the module
    /// </summary>
    public List<string> Procedures { get; set; } = [];
}

/// <summary>
/// Information about a VBA script
/// </summary>
public class ScriptInfo
{
    /// <summary>
    /// Name of the script module
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Type of the script module
    /// </summary>
    public string Type { get; set; } = string.Empty;

    /// <summary>
    /// Number of lines in the module
    /// </summary>
    public int LineCount { get; set; }

    /// <summary>
    /// List of procedures in the module
    /// </summary>
    public List<string> Procedures { get; set; } = [];
}

/// <summary>
/// File validation details for FileCommands.Test
/// </summary>
public class FileValidationInfo
{
    /// <summary>
    /// Full file path being validated
    /// </summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>
    /// Whether the file exists
    /// </summary>
    public bool Exists { get; set; }

    /// <summary>
    /// Size of the file in bytes
    /// </summary>
    public long Size { get; set; }

    /// <summary>
    /// File extension
    /// </summary>
    public string Extension { get; set; } = string.Empty;

    /// <summary>
    /// Last modification time
    /// </summary>
    public DateTime LastModified { get; set; }

    /// <summary>
    /// Whether the file is a valid Excel workbook
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// Optional message describing validation outcome (missing file, invalid extension, etc.)
    /// </summary>
    public string? Message { get; set; }
}

/// <summary>
/// Result for cell operations
/// </summary>
public class CellValueResult : ResultBase
{
    /// <summary>
    /// Address of the cell (e.g., A1)
    /// </summary>
    public string CellAddress { get; set; } = string.Empty;

    /// <summary>
    /// Current value of the cell
    /// </summary>
    public object? Value { get; set; }

    /// <summary>
    /// Type of the value
    /// </summary>
    public string ValueType { get; set; } = string.Empty;

    /// <summary>
    /// Formula in the cell, if any
    /// </summary>
    public string? Formula { get; set; }
}

/// <summary>
/// Result for Excel range value operations
/// </summary>
/// <remarks>
/// Property names: sn=SheetName, ra=RangeAddress, d=Values (data), r=RowCount, c=ColumnCount
/// </remarks>
public class RangeValueResult : ResultBase
{
    /// <summary>
    /// Sheet name
    /// </summary>
    [JsonPropertyName("sn")]
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Range address (e.g., A1:D10)
    /// </summary>
    [JsonPropertyName("ra")]
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>
    /// 2D array of cell values (row-major order)
    /// </summary>
    [JsonPropertyName("d")]
    public List<List<object?>> Values { get; set; } = [];

    /// <summary>
    /// Number of rows in the range
    /// </summary>
    [JsonPropertyName("r")]
    public int RowCount { get; set; }

    /// <summary>
    /// Number of columns in the range
    /// </summary>
    [JsonPropertyName("c")]
    public int ColumnCount { get; set; }
}

/// <summary>
/// Result for Excel range formula operations
/// </summary>
/// <remarks>
/// Property names: sn=SheetName, ra=RangeAddress, f=Formulas, d=Values, r=RowCount, c=ColumnCount
/// </remarks>
public class RangeFormulaResult : ResultBase
{
    /// <summary>
    /// Sheet name
    /// </summary>
    [JsonPropertyName("sn")]
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Range address (e.g., A1:D10)
    /// </summary>
    [JsonPropertyName("ra")]
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>
    /// 2D array of cell formulas (row-major order, empty string if no formula)
    /// </summary>
    [JsonPropertyName("f")]
    public List<List<string>> Formulas { get; set; } = [];

    /// <summary>
    /// 2D array of cell values (calculated results)
    /// </summary>
    [JsonPropertyName("d")]
    public List<List<object?>> Values { get; set; } = [];

    /// <summary>
    /// Number of rows in the range
    /// </summary>
    [JsonPropertyName("r")]
    public int RowCount { get; set; }

    /// <summary>
    /// Number of columns in the range
    /// </summary>
    [JsonPropertyName("c")]
    public int ColumnCount { get; set; }
}

/// <summary>
/// Result for range find operations
/// </summary>
/// <remarks>
/// Property names: sn=SheetName, ra=RangeAddress
/// </remarks>
public class RangeFindResult : ResultBase
{
    /// <summary>
    /// Sheet name
    /// </summary>
    [JsonPropertyName("sn")]
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Range address that was searched
    /// </summary>
    [JsonPropertyName("ra")]
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>
    /// Search value
    /// </summary>
    public string SearchValue { get; set; } = string.Empty;

    /// <summary>
    /// List of matching cells
    /// </summary>
    public List<RangeCell> MatchingCells { get; set; } = [];
}

/// <summary>
/// Represents a single cell in a range
/// </summary>
public class RangeCell
{
    /// <summary>
    /// Cell address (e.g., "A5")
    /// </summary>
    public string Address { get; set; } = string.Empty;

    /// <summary>
    /// Row number (1-based)
    /// </summary>
    public int Row { get; set; }

    /// <summary>
    /// Column number (1-based)
    /// </summary>
    public int Column { get; set; }

    /// <summary>
    /// Cell value
    /// </summary>
    public object? Value { get; set; }
}

/// <summary>
/// Result for range information operations
/// </summary>
public class RangeInfoResult : ResultBase
{
    /// <summary>
    /// Sheet name
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Absolute address from Excel COM (e.g., "$A$1:$D$10")
    /// </summary>
    public string Address { get; set; } = string.Empty;

    /// <summary>
    /// Number of rows (Excel COM: range.Rows.Count)
    /// </summary>
    public int RowCount { get; set; }

    /// <summary>
    /// Number of columns (Excel COM: range.Columns.Count)
    /// </summary>
    public int ColumnCount { get; set; }

    /// <summary>
    /// Number format code (Excel COM: range.NumberFormat, first cell)
    /// </summary>
    public string? NumberFormat { get; set; }
}

/// <summary>
/// Result for hyperlink operations
/// </summary>
public class RangeHyperlinkResult : ResultBase
{
    /// <summary>
    /// Sheet name
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Range or cell address
    /// </summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>
    /// List of hyperlinks
    /// </summary>
    public List<HyperlinkInfo> Hyperlinks { get; set; } = [];
}

/// <summary>
/// Result for Excel range style operations
/// </summary>
public class RangeStyleResult : ResultBase
{
    /// <summary>
    /// Sheet name
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Range address (e.g., A1:D10)
    /// </summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>
    /// Current style name applied to the range (first cell)
    /// </summary>
    public string StyleName { get; set; } = string.Empty;

    /// <summary>
    /// Whether this is a built-in Excel style
    /// </summary>
    public bool IsBuiltInStyle { get; set; }

    /// <summary>
    /// Additional style information if available
    /// </summary>
    public string? StyleDescription { get; set; }
}

/// <summary>
/// Result for VBA trust operations
/// </summary>
public class VbaTrustResult : ResultBase
{
    /// <summary>
    /// Whether VBA project access is trusted
    /// </summary>
    public bool IsTrusted { get; set; }

    /// <summary>
    /// Number of VBA components found (when checking trust)
    /// </summary>
    public int ComponentCount { get; set; }

    /// <summary>
    /// Registry paths where trust was set
    /// </summary>
    public List<string> RegistryPathsSet { get; set; } = [];

    /// <summary>
    /// Manual setup instructions if automated setup failed
    /// </summary>
    public string? ManualInstructions { get; set; }
}

/// <summary>
/// Power Query privacy level options for data combining
/// OBSOLETE: Privacy levels cannot be set programmatically.
/// Configure manually in Excel: File → Options → Privacy
/// </summary>
[Obsolete("Privacy levels not supported. Configure manually in Excel UI: File → Options → Privacy")]
public enum PowerQueryPrivacyLevel
{
    /// <summary>
    /// Ignores privacy levels, allows combining any data sources (least secure)
    /// </summary>
    None,

    /// <summary>
    /// Prevents sharing data with other sources (most secure, recommended for sensitive data)
    /// </summary>
    Private,

    /// <summary>
    /// Data can be shared within organization (recommended for internal data)
    /// </summary>
    Organizational,

    /// <summary>
    /// Publicly available data sources (appropriate for public APIs)
    /// </summary>
    Public
}

/// <summary>
/// Result for Power Query refresh operations with error detection
/// </summary>
public class PowerQueryRefreshResult : ResultBase
{
    /// <summary>
    /// Name of the query that was refreshed
    /// </summary>
    public string QueryName { get; set; } = string.Empty;

    /// <summary>
    /// Whether query has errors after refresh attempt
    /// </summary>
    public bool HasErrors { get; set; }

    /// <summary>
    /// List of error messages detected
    /// </summary>
    public List<string> ErrorMessages { get; set; } = [];

    /// <summary>
    /// When the refresh was attempted
    /// </summary>
    public DateTime RefreshTime { get; set; }

    /// <summary>
    /// Whether this is a connection-only query
    /// </summary>
    public bool IsConnectionOnly { get; set; }

    /// <summary>
    /// Worksheet name where data was loaded (if applicable)
    /// </summary>
    public string? LoadedToSheet { get; set; }
}

/// <summary>
/// Result for Power Query error checking
/// </summary>
public class PowerQueryErrorCheckResult : ResultBase
{
    /// <summary>
    /// Name of the query checked
    /// </summary>
    public string QueryName { get; set; } = string.Empty;

    /// <summary>
    /// Whether errors were detected
    /// </summary>
    public bool HasErrors { get; set; }

    /// <summary>
    /// List of error messages
    /// </summary>
    public List<string> ErrorMessages { get; set; } = [];

    /// <summary>
    /// Category of error (Authentication, Connectivity, Privacy, Syntax, Permissions, Unknown)
    /// </summary>
    public string? ErrorCategory { get; set; }

    /// <summary>
    /// Whether this is a connection-only query
    /// </summary>
    public bool IsConnectionOnly { get; set; }

    /// <summary>
    /// Additional message
    /// </summary>
    public string? Message { get; set; }
}

#region Connection Result Types

/// <summary>
/// Result for listing connections in a workbook
/// </summary>
public class ConnectionListResult : ResultBase
{
    /// <summary>
    /// List of connections in the workbook
    /// </summary>
    public List<ConnectionInfo> Connections { get; set; } = [];
}

/// <summary>
/// Information about a connection
/// </summary>
public class ConnectionInfo
{
    /// <summary>
    /// Connection name
    /// </summary>
    public string Name { get; init; } = "";

    /// <summary>
    /// Connection description
    /// </summary>
    public string? Description { get; init; }

    /// <summary>
    /// Connection type (OLEDB, ODBC, XML, Text, Web, DataFeed, Model, Worksheet, NoSource)
    /// </summary>
    public string Type { get; init; } = "";

    /// <summary>
    /// Last refresh date/time (if available)
    /// </summary>
    public DateTime? LastRefresh { get; init; }

    /// <summary>
    /// Whether the connection refreshes in background
    /// </summary>
    public bool BackgroundQuery { get; init; }

    /// <summary>
    /// Whether the connection refreshes when file opens
    /// </summary>
    public bool RefreshOnFileOpen { get; init; }

    /// <summary>
    /// Whether this is a Power Query connection
    /// </summary>
    public bool IsPowerQuery { get; init; }
}

/// <summary>
/// Result for viewing connection details
/// </summary>
public class ConnectionViewResult : ResultBase
{
    /// <summary>
    /// Connection name
    /// </summary>
    public string ConnectionName { get; set; } = "";

    /// <summary>
    /// Connection type (OLEDB, ODBC, XML, Text, Web, DataFeed, Model, Worksheet, NoSource)
    /// </summary>
    public string Type { get; set; } = "";

    /// <summary>
    /// Connection string (SANITIZED - passwords masked)
    /// </summary>
    public string ConnectionString { get; set; } = "";

    /// <summary>
    /// Command text (SQL query, M code reference, etc.)
    /// </summary>
    public string? CommandText { get; set; }

    /// <summary>
    /// Command type (SQL, Table, Default, etc.)
    /// </summary>
    public string? CommandType { get; set; }

    /// <summary>
    /// Whether this is a Power Query connection
    /// </summary>
    public bool IsPowerQuery { get; set; }

    /// <summary>
    /// Full connection definition as JSON
    /// </summary>
    public string DefinitionJson { get; set; } = "";
}

/// <summary>
/// Result for getting connection properties
/// </summary>
public class ConnectionPropertiesResult : ResultBase
{
    /// <summary>
    /// Connection name
    /// </summary>
    public string ConnectionName { get; set; } = "";

    /// <summary>
    /// Whether the connection refreshes in background
    /// </summary>
    public bool BackgroundQuery { get; set; }

    /// <summary>
    /// Whether the connection refreshes when file opens
    /// </summary>
    public bool RefreshOnFileOpen { get; set; }

    /// <summary>
    /// Whether password is saved with connection
    /// </summary>
    public bool SavePassword { get; set; }

    /// <summary>
    /// Refresh period in minutes (0 = no automatic refresh)
    /// </summary>
    public int RefreshPeriod { get; set; }
}

#endregion

#region Data Model Result Types

/// <summary>
/// Result for listing Data Model tables
/// </summary>
public class DataModelTableListResult : ResultBase
{
    /// <summary>
    /// List of tables in the Data Model
    /// </summary>
    public List<DataModelTableInfo> Tables { get; set; } = [];
}

/// <summary>
/// Information about a Data Model table
/// </summary>
public class DataModelTableInfo
{
    /// <summary>
    /// Table name
    /// </summary>
    public string Name { get; init; } = "";

    /// <summary>
    /// Source query or connection name
    /// </summary>
    public string SourceName { get; init; } = "";

    /// <summary>
    /// Number of rows in the table
    /// </summary>
    public int RecordCount { get; init; }
}

/// <summary>
/// Result for listing DAX measures
/// </summary>
public class DataModelMeasureListResult : ResultBase
{
    /// <summary>
    /// List of DAX measures in the model
    /// </summary>
    public List<DataModelMeasureInfo> Measures { get; set; } = [];
}

/// <summary>
/// Information about a DAX measure
/// </summary>
public class DataModelMeasureInfo
{
    /// <summary>
    /// Measure name
    /// </summary>
    public string Name { get; init; } = "";

    /// <summary>
    /// Table name where measure is defined
    /// </summary>
    public string Table { get; init; } = "";

    /// <summary>
    /// DAX formula preview (truncated for display)
    /// </summary>
    public string FormulaPreview { get; init; } = "";

    /// <summary>
    /// Measure description (if available)
    /// </summary>
    public string? Description { get; init; }
}

/// <summary>
/// Format information for a Data Model measure.
/// Represents the polymorphic ModelFormat* COM objects as structured JSON.
/// </summary>
public class MeasureFormatInfo
{
    /// <summary>
    /// Format type: General, Currency, Decimal, Percentage, WholeNumber, Scientific, Boolean, Date
    /// </summary>
    public string Type { get; set; } = "General";

    /// <summary>
    /// Currency symbol (e.g., "$", "€", "£"). Only present for Currency format.
    /// </summary>
    public string? Symbol { get; set; }

    /// <summary>
    /// Number of decimal places. Present for Currency, Decimal, Percentage formats.
    /// </summary>
    public int? DecimalPlaces { get; set; }

    /// <summary>
    /// Whether to use thousand separator (e.g., 1,000 vs 1000). Present for numeric formats.
    /// </summary>
    public bool? UseThousandSeparator { get; set; }
}

/// <summary>
/// Result for viewing measure details
/// </summary>
public class DataModelMeasureViewResult : ResultBase
{
    /// <summary>
    /// Measure name
    /// </summary>
    public string MeasureName { get; set; } = "";

    /// <summary>
    /// Table name where measure is defined
    /// </summary>
    public string TableName { get; set; } = "";

    /// <summary>
    /// Complete DAX formula
    /// </summary>
    public string DaxFormula { get; set; } = "";

    /// <summary>
    /// Measure description
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Format information extracted from ModelFormat* COM objects.
    /// Contains Type, Symbol, DecimalPlaces, UseThousandSeparator as applicable.
    /// </summary>
    public MeasureFormatInfo? FormatInfo { get; set; }

    /// <summary>
    /// Number of characters in DAX formula
    /// </summary>
    public int CharacterCount { get; set; }
}

/// <summary>
/// Result for listing model relationships
/// </summary>
public class DataModelRelationshipListResult : ResultBase
{
    /// <summary>
    /// List of relationships in the model
    /// </summary>
    public List<DataModelRelationshipInfo> Relationships { get; set; } = [];
}

/// <summary>
/// Result for reading a single relationship
/// </summary>
public class DataModelRelationshipViewResult : ResultBase
{
    /// <summary>
    /// The relationship details
    /// </summary>
    public DataModelRelationshipInfo? Relationship { get; set; }
}

/// <summary>
/// Information about a table relationship
/// </summary>
public class DataModelRelationshipInfo
{
    /// <summary>
    /// Source table name (foreign key side)
    /// </summary>
    public string FromTable { get; init; } = "";

    /// <summary>
    /// Source column name (foreign key)
    /// </summary>
    public string FromColumn { get; init; } = "";

    /// <summary>
    /// Target table name (primary key side)
    /// </summary>
    public string ToTable { get; init; } = "";

    /// <summary>
    /// Target column name (primary key)
    /// </summary>
    public string ToColumn { get; init; } = "";

    /// <summary>
    /// Whether this relationship is active
    /// </summary>
    public bool IsActive { get; init; }
}

/// <summary>
/// Result for DAX formula validation
/// </summary>
public class DataModelValidationResult : ResultBase
{
    /// <summary>
    /// Whether the DAX formula is valid
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// Validation error message (if not valid)
    /// </summary>
    public string? ValidationError { get; set; }

    /// <summary>
    /// DAX formula that was validated
    /// </summary>
    public string DaxFormula { get; set; } = "";
}

/// <summary>
/// Result for listing calculated columns
/// </summary>
public class DataModelCalculatedColumnListResult : ResultBase
{
    /// <summary>
    /// List of calculated columns in the model
    /// </summary>
    public List<DataModelCalculatedColumnInfo> CalculatedColumns { get; set; } = [];
}

/// <summary>
/// Information about a calculated column
/// </summary>
public class DataModelCalculatedColumnInfo
{
    /// <summary>
    /// Column name
    /// </summary>
    public string Name { get; init; } = "";

    /// <summary>
    /// Table name where column is defined
    /// </summary>
    public string Table { get; init; } = "";

    /// <summary>
    /// DAX formula preview (truncated for display)
    /// </summary>
    public string FormulaPreview { get; init; } = "";

    /// <summary>
    /// Data type of the column
    /// </summary>
    public string DataType { get; init; } = "";

    /// <summary>
    /// Column description (if available)
    /// </summary>
    public string? Description { get; init; }
}

/// <summary>
/// Result for viewing calculated column details
/// </summary>
public class DataModelCalculatedColumnViewResult : ResultBase
{
    /// <summary>
    /// Column name
    /// </summary>
    public string ColumnName { get; set; } = "";

    /// <summary>
    /// Table name where column is defined
    /// </summary>
    public string TableName { get; set; } = "";

    /// <summary>
    /// Complete DAX formula
    /// </summary>
    public string DaxFormula { get; set; } = "";

    /// <summary>
    /// Column description
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Data type of the column
    /// </summary>
    public string DataType { get; set; } = "";

    /// <summary>
    /// Number of characters in DAX formula
    /// </summary>
    public int CharacterCount { get; set; }
}

#endregion

#region Table (ListObject) Results

/// <summary>
/// Result for listing Excel Tables
/// </summary>
/// <remarks>
/// Property names: ts=Tables
/// </remarks>
public class TableListResult : ResultBase
{
    /// <summary>
    /// List of Excel Tables in the workbook
    /// </summary>
    [JsonPropertyName("ts")]
    public List<TableInfo> Tables { get; set; } = [];
}

/// <summary>
/// Result for getting detailed information about an Excel Table
/// </summary>
/// <remarks>
/// Property names: t=Table
/// </remarks>
public class TableInfoResult : ResultBase
{
    /// <summary>
    /// Detailed information about the Excel Table
    /// </summary>
    [JsonPropertyName("t")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public TableInfo? Table { get; set; }
}

/// <summary>
/// Information about an Excel Table (ListObject)
/// </summary>
/// <remarks>
/// Property names: n=Name, sn=SheetName, ra=Range, hh=HasHeaders, st=TableStyle, r=RowCount, c=ColumnCount, cols=Columns, tot=ShowTotals
/// </remarks>
public class TableInfo
{
    /// <summary>
    /// Name of the table
    /// </summary>
    [JsonPropertyName("n")]
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Worksheet containing the table
    /// </summary>
    [JsonPropertyName("sn")]
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Range address of the table (e.g., "A1:D10")
    /// </summary>
    [JsonPropertyName("ra")]
    public string Range { get; set; } = string.Empty;

    /// <summary>
    /// Whether the table has headers
    /// </summary>
    [JsonPropertyName("hh")]
    public bool HasHeaders { get; set; } = true;

    /// <summary>
    /// Table style name (e.g., "TableStyleMedium2")
    /// </summary>
    [JsonPropertyName("st")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? TableStyle { get; set; }

    /// <summary>
    /// Number of rows (excluding header)
    /// </summary>
    [JsonPropertyName("r")]
    public int RowCount { get; set; }

    /// <summary>
    /// Number of columns
    /// </summary>
    [JsonPropertyName("c")]
    public int ColumnCount { get; set; }

    /// <summary>
    /// Column names (if table has headers)
    /// </summary>
    [JsonPropertyName("cols")]
    public List<string> Columns { get; set; } = [];

    /// <summary>
    /// Whether the table has a total row
    /// </summary>
    [JsonPropertyName("tot")]
    public bool ShowTotals { get; set; }
}

/// <summary>
/// Result for reading Excel Table data
/// </summary>
/// <remarks>
/// Property names: tn=TableName, h=Headers, d=Data, r=RowCount, c=ColumnCount
/// </remarks>
public class TableDataResult : ResultBase
{
    /// <summary>
    /// Name of the table
    /// </summary>
    [JsonPropertyName("tn")]
    public string TableName { get; set; } = string.Empty;

    /// <summary>
    /// Column headers
    /// </summary>
    [JsonPropertyName("h")]
    public List<string> Headers { get; set; } = [];

    /// <summary>
    /// Data rows (each row is a list of cell values)
    /// </summary>
    [JsonPropertyName("d")]
    public List<List<object?>> Data { get; set; } = [];

    /// <summary>
    /// Number of rows (excluding header)
    /// </summary>
    [JsonPropertyName("r")]
    public int RowCount { get; set; }

    /// <summary>
    /// Number of columns
    /// </summary>
    public int ColumnCount { get; set; }
}

/// <summary>
/// Result for getting filter state of an Excel Table
/// </summary>
public class TableFilterResult : ResultBase
{
    /// <summary>
    /// Name of the table
    /// </summary>
    public string TableName { get; set; } = string.Empty;

    /// <summary>
    /// Filter information for each column
    /// </summary>
    public List<ColumnFilter> ColumnFilters { get; set; } = [];

    /// <summary>
    /// Whether any filters are active
    /// </summary>
    public bool HasActiveFilters { get; set; }
}

/// <summary>
/// Filter information for a table column
/// </summary>
public class ColumnFilter
{
    /// <summary>
    /// Column name
    /// </summary>
    public string ColumnName { get; set; } = string.Empty;

    /// <summary>
    /// Column index (1-based)
    /// </summary>
    public int ColumnIndex { get; set; }

    /// <summary>
    /// Whether this column has an active filter
    /// </summary>
    public bool IsFiltered { get; set; }

    /// <summary>
    /// Filter criteria (if single criteria)
    /// </summary>
    public string? Criteria { get; set; }

    /// <summary>
    /// Filter values (if multiple values)
    /// </summary>
    public List<string>? FilterValues { get; set; }
}

/// <summary>
/// Excel Table regions for structured references
/// </summary>
public enum TableRegion
{
    /// <summary>
    /// Entire table including headers, data, and totals (TableName[#All])
    /// </summary>
    All,

    /// <summary>
    /// Data rows only, excluding headers and totals (TableName[#Data])
    /// </summary>
    Data,

    /// <summary>
    /// Header row only (TableName[#Headers])
    /// </summary>
    Headers,

    /// <summary>
    /// Totals row only (TableName[#Totals])
    /// </summary>
    Totals,

    /// <summary>
    /// This row in formula context (TableName[@])
    /// </summary>
    ThisRow
}

/// <summary>
/// Result for getting structured reference information for a table region
/// </summary>
public class TableStructuredReferenceResult : ResultBase
{
    /// <summary>
    /// Name of the table
    /// </summary>
    public string TableName { get; set; } = string.Empty;

    /// <summary>
    /// Table region requested
    /// </summary>
    public TableRegion Region { get; set; }

    /// <summary>
    /// Excel range address for the region (e.g., "$A$1:$D$100")
    /// </summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>
    /// Structured reference formula (e.g., "SalesTable[#Data]")
    /// </summary>
    public string StructuredReference { get; set; } = string.Empty;

    /// <summary>
    /// Sheet name where the table is located
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Column name (if requesting specific column reference)
    /// </summary>
    public string? ColumnName { get; set; }

    /// <summary>
    /// Number of rows in the region
    /// </summary>
    public int RowCount { get; set; }

    /// <summary>
    /// Number of columns in the region
    /// </summary>
    public int ColumnCount { get; set; }
}

/// <summary>
/// Sort column specification for table sorting
/// </summary>
public class TableSortColumn
{
    /// <summary>
    /// Column name to sort by
    /// </summary>
    public string ColumnName { get; set; } = string.Empty;

    /// <summary>
    /// Whether to sort in ascending order (true) or descending (false)
    /// </summary>
    public bool Ascending { get; set; } = true;
}

#endregion

#region Hyperlink Results

/// <summary>
/// Information about a hyperlink in an Excel cell
/// </summary>
public class HyperlinkInfo
{
    /// <summary>
    /// Cell address containing the hyperlink
    /// </summary>
    public string CellAddress { get; set; } = string.Empty;

    /// <summary>
    /// Hyperlink URL or file path
    /// </summary>
    public string Address { get; set; } = string.Empty;

    /// <summary>
    /// Sub-address within the target (e.g., sheet reference)
    /// </summary>
    public string? SubAddress { get; set; }

    /// <summary>
    /// Display text (visible text in cell)
    /// </summary>
    public string DisplayText { get; set; } = string.Empty;

    /// <summary>
    /// Tooltip/ScreenTip text
    /// </summary>
    public string? ScreenTip { get; set; }

    /// <summary>
    /// Whether the hyperlink points to another location in the workbook
    /// </summary>
    public bool IsInternal { get; set; }
}

/// <summary>
/// Result for listing hyperlinks in a worksheet
/// </summary>
public class HyperlinkListResult : ResultBase
{
    /// <summary>
    /// List of hyperlinks in the worksheet
    /// </summary>
    public List<HyperlinkInfo> Hyperlinks { get; set; } = [];

    /// <summary>
    /// Total count of hyperlinks
    /// </summary>
    public int Count { get; set; }

    /// <summary>
    /// Sheet name
    /// </summary>
    public string SheetName { get; set; } = string.Empty;
}

/// <summary>
/// Result for getting hyperlink information from a specific cell
/// </summary>
public class HyperlinkInfoResult : ResultBase
{
    /// <summary>
    /// Hyperlink information (null if no hyperlink exists)
    /// </summary>
    public HyperlinkInfo? Hyperlink { get; set; }

    /// <summary>
    /// Whether a hyperlink exists at the specified cell
    /// </summary>
    public bool HasHyperlink { get; set; }

    /// <summary>
    /// Sheet name
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Cell address
    /// </summary>
    public string CellAddress { get; set; } = string.Empty;
}

#endregion

#region Number Formatting Results

/// <summary>
/// Result for range number format operations
/// </summary>
public class RangeNumberFormatResult : ResultBase
{
    /// <summary>
    /// Sheet name
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Range address (e.g., A1:D10)
    /// </summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>
    /// 2D array of number format codes (matches range dimensions)
    /// </summary>
    public List<List<string>> Formats { get; set; } = [];

    /// <summary>
    /// Number of rows in the range
    /// </summary>
    public int RowCount { get; set; }

    /// <summary>
    /// Number of columns in the range
    /// </summary>
    public int ColumnCount { get; set; }
}

#endregion

#region Validation Results

/// <summary>
/// Result for range validation operations
/// </summary>
public class RangeValidationResult : ResultBase
{
    /// <summary>
    /// Sheet name
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Range address
    /// </summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>
    /// Whether the range has validation
    /// </summary>
    public bool HasValidation { get; set; }

    /// <summary>
    /// Validation type (list, whole, decimal, date, time, textlength, custom)
    /// </summary>
    public string? ValidationType { get; set; }

    /// <summary>
    /// Validation operator (between, equal, greaterthan, etc.)
    /// </summary>
    public string? ValidationOperator { get; set; }

    /// <summary>
    /// First formula/value
    /// </summary>
    public string? Formula1 { get; set; }

    /// <summary>
    /// Second formula/value (for Between operator)
    /// </summary>
    public string? Formula2 { get; set; }

    /// <summary>
    /// Whether to ignore blank cells
    /// </summary>
    public bool IgnoreBlank { get; set; }

    /// <summary>
    /// Whether to show input message
    /// </summary>
    public bool ShowInputMessage { get; set; }

    /// <summary>
    /// Input message title
    /// </summary>
    public string? InputTitle { get; set; }

    /// <summary>
    /// Input message text
    /// </summary>
    public string? InputMessage { get; set; }

    /// <summary>
    /// Whether to show error alert
    /// </summary>
    public bool ShowErrorAlert { get; set; }

    /// <summary>
    /// Error alert style (stop, warning, information)
    /// </summary>
    public string? ErrorStyle { get; set; }

    /// <summary>
    /// Error alert title
    /// </summary>
    public string? ErrorTitle { get; set; }

    /// <summary>
    /// Error alert message text
    /// </summary>
    public string? ValidationErrorMessage { get; set; }
}

#endregion

#region Cell Merge and Protection Results

/// <summary>
/// Result for range merge information
/// </summary>
public class RangeMergeInfoResult : ResultBase
{
    /// <summary>
    /// Sheet name
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Range address
    /// </summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>
    /// Whether the range contains merged cells
    /// </summary>
    public bool IsMerged { get; set; }
}

/// <summary>
/// Result for cell lock information
/// </summary>
public class RangeLockInfoResult : ResultBase
{
    /// <summary>
    /// Sheet name
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Range address
    /// </summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>
    /// Whether the cells are locked
    /// </summary>
    public bool IsLocked { get; set; }
}

#endregion
