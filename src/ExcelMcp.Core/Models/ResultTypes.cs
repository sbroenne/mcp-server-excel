using System.Collections.Generic;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Base result type for all Core operations
/// </summary>
public abstract class ResultBase
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public string? FilePath { get; set; }
}

/// <summary>
/// Result for operations that don't return data (create, delete, etc.)
/// </summary>
public class OperationResult : ResultBase
{
    public string? Action { get; set; }
}

/// <summary>
/// Result for listing worksheets
/// </summary>
public class WorksheetListResult : ResultBase
{
    public List<WorksheetInfo> Worksheets { get; set; } = new();
}

public class WorksheetInfo
{
    public string Name { get; set; } = string.Empty;
    public int Index { get; set; }
    public bool Visible { get; set; }
}

/// <summary>
/// Result for reading worksheet data
/// </summary>
public class WorksheetDataResult : ResultBase
{
    public string SheetName { get; set; } = string.Empty;
    public string Range { get; set; } = string.Empty;
    public List<List<object?>> Data { get; set; } = new();
    public List<string> Headers { get; set; } = new();
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
}

/// <summary>
/// Result for listing Power Queries
/// </summary>
public class PowerQueryListResult : ResultBase
{
    public List<PowerQueryInfo> Queries { get; set; } = new();
}

public class PowerQueryInfo
{
    public string Name { get; set; } = string.Empty;
    public string Formula { get; set; } = string.Empty;
    public string FormulaPreview { get; set; } = string.Empty;
    public bool IsConnectionOnly { get; set; }
}

/// <summary>
/// Result for viewing Power Query code
/// </summary>
public class PowerQueryViewResult : ResultBase
{
    public string QueryName { get; set; } = string.Empty;
    public string MCode { get; set; } = string.Empty;
    public int CharacterCount { get; set; }
    public bool IsConnectionOnly { get; set; }
}

/// <summary>
/// Result for listing named ranges/parameters
/// </summary>
public class ParameterListResult : ResultBase
{
    public List<ParameterInfo> Parameters { get; set; } = new();
}

public class ParameterInfo
{
    public string Name { get; set; } = string.Empty;
    public string RefersTo { get; set; } = string.Empty;
    public object? Value { get; set; }
    public string ValueType { get; set; } = string.Empty;
}

/// <summary>
/// Result for getting parameter value
/// </summary>
public class ParameterValueResult : ResultBase
{
    public string ParameterName { get; set; } = string.Empty;
    public object? Value { get; set; }
    public string ValueType { get; set; } = string.Empty;
    public string RefersTo { get; set; } = string.Empty;
}

/// <summary>
/// Result for listing VBA scripts
/// </summary>
public class ScriptListResult : ResultBase
{
    public List<ScriptInfo> Scripts { get; set; } = new();
}

public class ScriptInfo
{
    public string Name { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
    public List<string> Procedures { get; set; } = new();
}

/// <summary>
/// Result for file operations
/// </summary>
public class FileValidationResult : ResultBase
{
    public bool Exists { get; set; }
    public long Size { get; set; }
    public string Extension { get; set; } = string.Empty;
    public DateTime LastModified { get; set; }
    public bool IsValid { get; set; }
}

/// <summary>
/// Result for cell operations
/// </summary>
public class CellValueResult : ResultBase
{
    public string CellAddress { get; set; } = string.Empty;
    public object? Value { get; set; }
    public string ValueType { get; set; } = string.Empty;
    public string? Formula { get; set; }
}