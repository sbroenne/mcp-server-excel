namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// QueryTable creation options
/// </summary>
public class QueryTableCreateOptions
{
    /// <summary>
    /// Whether to run the query in the background (asynchronously)
    /// </summary>
    public bool BackgroundQuery { get; init; }

    /// <summary>
    /// Whether to refresh the QueryTable when the file is opened
    /// </summary>
    public bool RefreshOnFileOpen { get; init; }

    /// <summary>
    /// Whether to save the password with the QueryTable
    /// </summary>
    public bool SavePassword { get; init; }

    /// <summary>
    /// Whether to preserve column information when refreshing
    /// </summary>
    public bool PreserveColumnInfo { get; init; } = true;

    /// <summary>
    /// Whether to preserve formatting when refreshing
    /// </summary>
    public bool PreserveFormatting { get; init; } = true;

    /// <summary>
    /// Whether to adjust column width automatically
    /// </summary>
    public bool AdjustColumnWidth { get; init; } = true;

    /// <summary>
    /// Whether to refresh immediately after creation (default: true for immediate feedback)
    /// </summary>
    public bool RefreshImmediately { get; init; } = true;
}

/// <summary>
/// QueryTable property update options
/// </summary>
public class QueryTableUpdateOptions
{
    /// <summary>
    /// Whether to run the query in the background (asynchronously)
    /// </summary>
    public bool? BackgroundQuery { get; init; }

    /// <summary>
    /// Whether to refresh the QueryTable when the file is opened
    /// </summary>
    public bool? RefreshOnFileOpen { get; init; }

    /// <summary>
    /// Whether to save the password with the QueryTable
    /// </summary>
    public bool? SavePassword { get; init; }

    /// <summary>
    /// Whether to preserve column information when refreshing
    /// </summary>
    public bool? PreserveColumnInfo { get; init; }

    /// <summary>
    /// Whether to preserve formatting when refreshing
    /// </summary>
    public bool? PreserveFormatting { get; init; }

    /// <summary>
    /// Whether to adjust column width automatically
    /// </summary>
    public bool? AdjustColumnWidth { get; init; }
}

/// <summary>
/// Result for QueryTable list operations
/// </summary>
public class QueryTableListResult : ResultBase
{
    /// <summary>
    /// List of QueryTables in the workbook
    /// </summary>
    public List<QueryTableInfo> QueryTables { get; set; } = [];
}

/// <summary>
/// Result for QueryTable info operations
/// </summary>
public class QueryTableInfoResult : ResultBase
{
    /// <summary>
    /// QueryTable information
    /// </summary>
    public QueryTableInfo? QueryTable { get; set; }
}

/// <summary>
/// Information about a QueryTable
/// </summary>
public class QueryTableInfo
{
    /// <summary>
    /// Name of the QueryTable
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Name of the worksheet containing the QueryTable
    /// </summary>
    public string WorksheetName { get; set; } = string.Empty;

    /// <summary>
    /// Range address of the QueryTable
    /// </summary>
    public string Range { get; set; } = string.Empty;

    /// <summary>
    /// Name of the associated connection (if any)
    /// </summary>
    public string ConnectionName { get; set; } = string.Empty;

    /// <summary>
    /// Connection string
    /// </summary>
    public string ConnectionString { get; set; } = string.Empty;

    /// <summary>
    /// Command text (SQL query, etc.)
    /// </summary>
    public string CommandText { get; set; } = string.Empty;

    /// <summary>
    /// Last refresh date/time
    /// </summary>
    public DateTime? LastRefresh { get; set; }

    /// <summary>
    /// Whether background query is enabled
    /// </summary>
    public bool BackgroundQuery { get; set; }

    /// <summary>
    /// Whether refresh on file open is enabled
    /// </summary>
    public bool RefreshOnFileOpen { get; set; }

    /// <summary>
    /// Whether column info is preserved
    /// </summary>
    public bool PreserveColumnInfo { get; set; }

    /// <summary>
    /// Whether formatting is preserved
    /// </summary>
    public bool PreserveFormatting { get; set; }

    /// <summary>
    /// Whether column width is adjusted automatically
    /// </summary>
    public bool AdjustColumnWidth { get; set; }

    /// <summary>
    /// Number of rows in the QueryTable
    /// </summary>
    public int RowCount { get; set; }

    /// <summary>
    /// Number of columns in the QueryTable
    /// </summary>
    public int ColumnCount { get; set; }
}
