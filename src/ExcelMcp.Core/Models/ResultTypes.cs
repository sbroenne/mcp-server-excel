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