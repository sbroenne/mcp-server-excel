using System.Collections.Generic;

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