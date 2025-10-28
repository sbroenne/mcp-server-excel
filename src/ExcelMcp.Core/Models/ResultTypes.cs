namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Base result type for all Core operations
/// </summary>
public abstract class ResultBase
{
    /// <summary>
    /// Indicates whether the operation was successful
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// Error message if operation failed
    /// </summary>
    public string? ErrorMessage { get; set; }

    /// <summary>
    /// File path of the Excel file
    /// </summary>
    public string? FilePath { get; set; }

    /// <summary>
    /// Suggested next actions for LLM workflow guidance
    /// </summary>
    public List<string> SuggestedNextActions { get; set; } = new();

    /// <summary>
    /// Contextual workflow hint for LLM
    /// </summary>
    public string? WorkflowHint { get; set; }
}

/// <summary>
/// Result for operations that don't return data (create, delete, etc.)
/// </summary>
public class OperationResult : ResultBase
{
    /// <summary>
    /// Action that was performed
    /// </summary>
    public string? Action { get; set; }
}

/// <summary>
/// Result for listing worksheets
/// </summary>
public class WorksheetListResult : ResultBase
{
    /// <summary>
    /// List of worksheets in the workbook
    /// </summary>
    public List<WorksheetInfo> Worksheets { get; set; } = new();
}

/// <summary>
/// Information about a worksheet
/// </summary>
public class WorksheetInfo
{
    /// <summary>
    /// Name of the worksheet
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Index of the worksheet (1-based)
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// Whether the worksheet is visible
    /// </summary>
    public bool Visible { get; set; }
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
    public List<List<object?>> Data { get; set; } = new();

    /// <summary>
    /// Column headers
    /// </summary>
    public List<string> Headers { get; set; } = new();

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
    public List<PowerQueryInfo> Queries { get; set; } = new();
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
/// Result for listing named ranges/parameters
/// </summary>
public class ParameterListResult : ResultBase
{
    /// <summary>
    /// List of named ranges/parameters
    /// </summary>
    public List<ParameterInfo> Parameters { get; set; } = new();
}

/// <summary>
/// Information about a named range/parameter
/// </summary>
public class ParameterInfo
{
    /// <summary>
    /// Name of the parameter
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// What the parameter refers to
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
public class ParameterValueResult : ResultBase
{
    /// <summary>
    /// Name of the parameter
    /// </summary>
    public string ParameterName { get; set; } = string.Empty;

    /// <summary>
    /// Current value
    /// </summary>
    public object? Value { get; set; }

    /// <summary>
    /// Type of the value
    /// </summary>
    public string ValueType { get; set; } = string.Empty;

    /// <summary>
    /// What the parameter refers to
    /// </summary>
    public string RefersTo { get; set; } = string.Empty;
}

/// <summary>
/// Result for listing VBA scripts
/// </summary>
public class ScriptListResult : ResultBase
{
    /// <summary>
    /// List of VBA scripts
    /// </summary>
    public List<ScriptInfo> Scripts { get; set; } = new();
}

/// <summary>
/// Result for viewing VBA module code
/// </summary>
public class ScriptViewResult : ResultBase
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
    public List<string> Procedures { get; set; } = new();
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
    public List<string> Procedures { get; set; } = new();
}

/// <summary>
/// Result for file operations
/// </summary>
public class FileValidationResult : ResultBase
{
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
    /// Whether the file is valid
    /// </summary>
    public bool IsValid { get; set; }
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
public class RangeValueResult : ResultBase
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
    /// 2D array of cell values (row-major order)
    /// </summary>
    public List<List<object?>> Values { get; set; } = new();

    /// <summary>
    /// Number of rows in the range
    /// </summary>
    public int RowCount { get; set; }

    /// <summary>
    /// Number of columns in the range
    /// </summary>
    public int ColumnCount { get; set; }
}

/// <summary>
/// Result for Excel range formula operations
/// </summary>
public class RangeFormulaResult : ResultBase
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
    /// 2D array of cell formulas (row-major order, empty string if no formula)
    /// </summary>
    public List<List<string>> Formulas { get; set; } = new();

    /// <summary>
    /// 2D array of cell values (calculated results)
    /// </summary>
    public List<List<object?>> Values { get; set; } = new();

    /// <summary>
    /// Number of rows in the range
    /// </summary>
    public int RowCount { get; set; }

    /// <summary>
    /// Number of columns in the range
    /// </summary>
    public int ColumnCount { get; set; }
}

/// <summary>
/// Result for range find operations
/// </summary>
public class RangeFindResult : ResultBase
{
    /// <summary>
    /// Sheet name
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Range address that was searched
    /// </summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>
    /// Search value
    /// </summary>
    public string SearchValue { get; set; } = string.Empty;

    /// <summary>
    /// List of matching cells
    /// </summary>
    public List<RangeCell> MatchingCells { get; set; } = new();
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
    public List<HyperlinkInfo> Hyperlinks { get; set; } = new();
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
    public List<string> RegistryPathsSet { get; set; } = new();

    /// <summary>
    /// Manual setup instructions if automated setup failed
    /// </summary>
    public string? ManualInstructions { get; set; }
}

/// <summary>
/// Power Query privacy level options for data combining
/// </summary>
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
/// Information about a query's detected privacy level
/// </summary>
public record QueryPrivacyInfo(string QueryName, PowerQueryPrivacyLevel PrivacyLevel);

/// <summary>
/// Result indicating Power Query operation requires privacy level specification
/// </summary>
public class PowerQueryPrivacyErrorResult : OperationResult
{
    /// <summary>
    /// Privacy levels detected in existing queries
    /// </summary>
    public List<QueryPrivacyInfo> ExistingPrivacyLevels { get; init; } = new();

    /// <summary>
    /// Recommended privacy level based on existing queries
    /// </summary>
    public PowerQueryPrivacyLevel RecommendedPrivacyLevel { get; init; }

    /// <summary>
    /// User-friendly explanation of the recommendation
    /// </summary>
    public string Explanation { get; init; } = "";

    /// <summary>
    /// Original error message from Excel
    /// </summary>
    public string OriginalError { get; init; } = "";
}

/// <summary>
/// Result indicating VBA operation requires trust access to VBA project object model.
/// Provides instructions for user to manually enable trust in Excel settings.
/// </summary>
public class VbaTrustRequiredResult : OperationResult
{
    /// <summary>
    /// Whether VBA trust is currently enabled
    /// </summary>
    public bool IsTrustEnabled { get; init; }

    /// <summary>
    /// Step-by-step instructions for enabling VBA trust
    /// </summary>
    public string[] SetupInstructions { get; init; } = new[]
    {
        "Open Excel",
        "Go to File → Options → Trust Center",
        "Click 'Trust Center Settings'",
        "Select 'Macro Settings'",
        "Check '✓ Trust access to the VBA project object model'",
        "Click OK twice to save settings"
    };

    /// <summary>
    /// Official Microsoft documentation URL
    /// </summary>
    public string DocumentationUrl { get; init; } = "https://support.microsoft.com/office/enable-or-disable-macros-in-office-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6";

    /// <summary>
    /// User-friendly explanation of why trust is required
    /// </summary>
    public string Explanation { get; init; } = "VBA operations require 'Trust access to the VBA project object model' to be enabled in Excel settings. This is a one-time setup that allows programmatic access to VBA code.";
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
    public List<string> ErrorMessages { get; set; } = new();

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
    public List<string> ErrorMessages { get; set; } = new();

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
    public List<ConnectionInfo> Connections { get; set; } = new();
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
    public List<DataModelTableInfo> Tables { get; set; } = new();
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

    /// <summary>
    /// Last refresh date/time (if available)
    /// </summary>
    public DateTime? RefreshDate { get; init; }
}

/// <summary>
/// Result for listing DAX measures
/// </summary>
public class DataModelMeasureListResult : ResultBase
{
    /// <summary>
    /// List of DAX measures in the model
    /// </summary>
    public List<DataModelMeasureInfo> Measures { get; set; } = new();
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
    /// Format string (e.g., "$#,##0.00", "0.00%")
    /// </summary>
    public string? FormatString { get; set; }

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
    public List<DataModelRelationshipInfo> Relationships { get; set; } = new();
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
    public List<DataModelCalculatedColumnInfo> CalculatedColumns { get; set; } = new();
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
public class TableListResult : ResultBase
{
    /// <summary>
    /// List of Excel Tables in the workbook
    /// </summary>
    public List<TableInfo> Tables { get; set; } = new();
}

/// <summary>
/// Result for getting detailed information about an Excel Table
/// </summary>
public class TableInfoResult : ResultBase
{
    /// <summary>
    /// Detailed information about the Excel Table
    /// </summary>
    public TableInfo? Table { get; set; }
}

/// <summary>
/// Information about an Excel Table (ListObject)
/// </summary>
public class TableInfo
{
    /// <summary>
    /// Name of the table
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Worksheet containing the table
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Range address of the table (e.g., "A1:D10")
    /// </summary>
    public string Range { get; set; } = string.Empty;

    /// <summary>
    /// Whether the table has headers
    /// </summary>
    public bool HasHeaders { get; set; } = true;

    /// <summary>
    /// Table style name (e.g., "TableStyleMedium2")
    /// </summary>
    public string? TableStyle { get; set; }

    /// <summary>
    /// Number of rows (excluding header)
    /// </summary>
    public int RowCount { get; set; }

    /// <summary>
    /// Number of columns
    /// </summary>
    public int ColumnCount { get; set; }

    /// <summary>
    /// Column names (if table has headers)
    /// </summary>
    public List<string> Columns { get; set; } = new();

    /// <summary>
    /// Whether the table has a total row
    /// </summary>
    public bool ShowTotals { get; set; }
}

/// <summary>
/// Result for reading Excel Table data
/// </summary>
public class TableDataResult : ResultBase
{
    /// <summary>
    /// Name of the table
    /// </summary>
    public string TableName { get; set; } = string.Empty;

    /// <summary>
    /// Column headers
    /// </summary>
    public List<string> Headers { get; set; } = new();

    /// <summary>
    /// Data rows (each row is a list of cell values)
    /// </summary>
    public List<List<object?>> Data { get; set; } = new();

    /// <summary>
    /// Number of rows (excluding header)
    /// </summary>
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
    public List<ColumnFilter> ColumnFilters { get; set; } = new();

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
    public List<HyperlinkInfo> Hyperlinks { get; set; } = new();

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
