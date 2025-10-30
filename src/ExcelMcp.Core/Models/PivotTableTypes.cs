namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// PivotTable field areas
/// </summary>
public enum PivotFieldArea
{
    /// <summary>
    /// Field is not displayed
    /// </summary>
    Hidden = 0,
    
    /// <summary>
    /// Field is in the Row area
    /// </summary>
    Row = 1,
    
    /// <summary>
    /// Field is in the Column area
    /// </summary>
    Column = 2,
    
    /// <summary>
    /// Field is in the Filter area (Page field)
    /// </summary>
    Filter = 3,
    
    /// <summary>
    /// Field is in the Values area (Data field)
    /// </summary>
    Value = 4
}

/// <summary>
/// Aggregation functions for PivotTable value fields
/// </summary>
public enum AggregationFunction
{
    /// <summary>
    /// Sum of values
    /// </summary>
    Sum,
    
    /// <summary>
    /// Count of all items
    /// </summary>
    Count,
    
    /// <summary>
    /// Average of values
    /// </summary>
    Average,
    
    /// <summary>
    /// Maximum value
    /// </summary>
    Max,
    
    /// <summary>
    /// Minimum value
    /// </summary>
    Min,
    
    /// <summary>
    /// Product of values
    /// </summary>
    Product,
    
    /// <summary>
    /// Count of numeric values
    /// </summary>
    CountNumbers,
    
    /// <summary>
    /// Standard deviation (sample)
    /// </summary>
    StdDev,
    
    /// <summary>
    /// Standard deviation (population)
    /// </summary>
    StdDevP,
    
    /// <summary>
    /// Variance (sample)
    /// </summary>
    Var,
    
    /// <summary>
    /// Variance (population)
    /// </summary>
    VarP
}

/// <summary>
/// Sort direction
/// </summary>
public enum SortDirection
{
    /// <summary>
    /// Ascending order
    /// </summary>
    Ascending,
    
    /// <summary>
    /// Descending order
    /// </summary>
    Descending
}

/// <summary>
/// Excel COM constants for PivotTable field orientation
/// </summary>
public static class XlPivotFieldOrientation
{
    /// <summary>
    /// Field not displayed
    /// </summary>
    public const int xlHidden = 0;
    
    /// <summary>
    /// Row area
    /// </summary>
    public const int xlRowField = 1;
    
    /// <summary>
    /// Column area
    /// </summary>
    public const int xlColumnField = 2;
    
    /// <summary>
    /// Filter area (Page field)
    /// </summary>
    public const int xlPageField = 3;
    
    /// <summary>
    /// Values area (Data field)
    /// </summary>
    public const int xlDataField = 4;
}

/// <summary>
/// Excel COM constants for consolidation functions
/// </summary>
public static class XlConsolidationFunction
{
    /// <summary>
    /// Sum function
    /// </summary>
    public const int xlSum = -4157;
    
    /// <summary>
    /// Count function
    /// </summary>
    public const int xlCount = -4112;
    
    /// <summary>
    /// Average function
    /// </summary>
    public const int xlAverage = -4106;
    
    /// <summary>
    /// Max function
    /// </summary>
    public const int xlMax = -4136;
    
    /// <summary>
    /// Min function
    /// </summary>
    public const int xlMin = -4139;
    
    /// <summary>
    /// Product function
    /// </summary>
    public const int xlProduct = -4149;
    
    /// <summary>
    /// Count numbers function
    /// </summary>
    public const int xlCountNums = -4113;
    
    /// <summary>
    /// Standard deviation function
    /// </summary>
    public const int xlStdDev = -4155;
    
    /// <summary>
    /// Standard deviation population function
    /// </summary>
    public const int xlStdDevP = -4156;
    
    /// <summary>
    /// Variance function
    /// </summary>
    public const int xlVar = -4164;
    
    /// <summary>
    /// Variance population function
    /// </summary>
    public const int xlVarP = -4165;
}

/// <summary>
/// Result for PivotTable creation operations
/// </summary>
public class PivotTableCreateResult : ResultBase
{
    /// <summary>
    /// Name of the created PivotTable
    /// </summary>
    public string PivotTableName { get; set; } = string.Empty;
    
    /// <summary>
    /// Sheet containing the PivotTable
    /// </summary>
    public string SheetName { get; set; } = string.Empty;
    
    /// <summary>
    /// Range occupied by the PivotTable
    /// </summary>
    public string Range { get; set; } = string.Empty;
    
    /// <summary>
    /// Source data reference
    /// </summary>
    public string SourceData { get; set; } = string.Empty;
    
    /// <summary>
    /// Number of rows in source data (excluding headers)
    /// </summary>
    public int SourceRowCount { get; set; }
    
    /// <summary>
    /// All available fields from source data
    /// </summary>
    public List<string> AvailableFields { get; set; } = new();
    
    /// <summary>
    /// Fields detected as numeric (suggested for Values area)
    /// </summary>
    public List<string> NumericFields { get; set; } = new();
    
    /// <summary>
    /// Fields detected as text (suggested for Rows/Columns/Filters)
    /// </summary>
    public List<string> TextFields { get; set; } = new();
    
    /// <summary>
    /// Fields detected as dates (suggested for grouping)
    /// </summary>
    public List<string> DateFields { get; set; } = new();
}

/// <summary>
/// Information about a PivotTable
/// </summary>
public class PivotTableInfo
{
    /// <summary>
    /// Name of the PivotTable
    /// </summary>
    public string Name { get; set; } = string.Empty;
    
    /// <summary>
    /// Sheet containing the PivotTable
    /// </summary>
    public string SheetName { get; set; } = string.Empty;
    
    /// <summary>
    /// Range occupied by the PivotTable
    /// </summary>
    public string Range { get; set; } = string.Empty;
    
    /// <summary>
    /// Source data reference
    /// </summary>
    public string SourceData { get; set; } = string.Empty;
    
    /// <summary>
    /// Number of row fields
    /// </summary>
    public int RowFieldCount { get; set; }
    
    /// <summary>
    /// Number of column fields
    /// </summary>
    public int ColumnFieldCount { get; set; }
    
    /// <summary>
    /// Number of value fields
    /// </summary>
    public int ValueFieldCount { get; set; }
    
    /// <summary>
    /// Number of filter fields
    /// </summary>
    public int FilterFieldCount { get; set; }
    
    /// <summary>
    /// Last refresh timestamp
    /// </summary>
    public DateTime? LastRefresh { get; set; }
}

/// <summary>
/// Result for listing PivotTables
/// </summary>
public class PivotTableListResult : ResultBase
{
    /// <summary>
    /// List of PivotTables in the workbook
    /// </summary>
    public List<PivotTableInfo> PivotTables { get; set; } = new();
}

/// <summary>
/// Result for getting PivotTable information
/// </summary>
public class PivotTableInfoResult : ResultBase
{
    /// <summary>
    /// Detailed information about the PivotTable
    /// </summary>
    public PivotTableInfo PivotTable { get; set; } = new();
    
    /// <summary>
    /// List of all fields in the PivotTable
    /// </summary>
    public List<PivotFieldInfo> Fields { get; set; } = new();
}

/// <summary>
/// Information about a PivotTable field
/// </summary>
public class PivotFieldInfo
{
    /// <summary>
    /// Name of the field
    /// </summary>
    public string Name { get; set; } = string.Empty;
    
    /// <summary>
    /// Custom name/caption
    /// </summary>
    public string CustomName { get; set; } = string.Empty;
    
    /// <summary>
    /// Area where the field is placed
    /// </summary>
    public PivotFieldArea Area { get; set; }
    
    /// <summary>
    /// Position within the area (1-based)
    /// </summary>
    public int Position { get; set; }
    
    /// <summary>
    /// Aggregation function (for value fields)
    /// </summary>
    public AggregationFunction? Function { get; set; }
    
    /// <summary>
    /// Data type of the field
    /// </summary>
    public string DataType { get; set; } = string.Empty;
}

/// <summary>
/// Result for field listing operations
/// </summary>
public class PivotFieldListResult : ResultBase
{
    /// <summary>
    /// List of all fields in the PivotTable
    /// </summary>
    public List<PivotFieldInfo> Fields { get; set; } = new();
}

/// <summary>
/// Result for field operations
/// </summary>
public class PivotFieldResult : ResultBase
{
    /// <summary>
    /// Name of the field
    /// </summary>
    public string FieldName { get; set; } = string.Empty;
    
    /// <summary>
    /// Custom name/caption
    /// </summary>
    public string CustomName { get; set; } = string.Empty;
    
    /// <summary>
    /// Area where the field is placed
    /// </summary>
    public PivotFieldArea Area { get; set; }
    
    /// <summary>
    /// Position within the area (1-based)
    /// </summary>
    public int Position { get; set; }
    
    /// <summary>
    /// Aggregation function (for value fields)
    /// </summary>
    public AggregationFunction? Function { get; set; }
    
    /// <summary>
    /// Number format
    /// </summary>
    public string? NumberFormat { get; set; }
    
    /// <summary>
    /// Available values for filtering
    /// </summary>
    public List<string> AvailableValues { get; set; } = new();
    
    /// <summary>
    /// Sample value for verification
    /// </summary>
    public object? SampleValue { get; set; }
    
    /// <summary>
    /// Data type of the field
    /// </summary>
    public string DataType { get; set; } = string.Empty;
}

/// <summary>
/// Result for PivotTable refresh operations
/// </summary>
public class PivotTableRefreshResult : ResultBase
{
    /// <summary>
    /// Name of the PivotTable
    /// </summary>
    public string PivotTableName { get; set; } = string.Empty;
    
    /// <summary>
    /// Refresh timestamp
    /// </summary>
    public DateTime RefreshTime { get; set; }
    
    /// <summary>
    /// Number of records in source data
    /// </summary>
    public int SourceRecordCount { get; set; }
    
    /// <summary>
    /// Previous record count (before refresh)
    /// </summary>
    public int PreviousRecordCount { get; set; }
    
    /// <summary>
    /// Whether structure changed
    /// </summary>
    public bool StructureChanged { get; set; }
    
    /// <summary>
    /// Fields added to source
    /// </summary>
    public List<string> NewFields { get; set; } = new();
    
    /// <summary>
    /// Fields removed from source
    /// </summary>
    public List<string> RemovedFields { get; set; } = new();
}

/// <summary>
/// Result for getting PivotTable data
/// </summary>
public class PivotTableDataResult : ResultBase
{
    /// <summary>
    /// Name of the PivotTable
    /// </summary>
    public string PivotTableName { get; set; } = string.Empty;
    
    /// <summary>
    /// 2D array of PivotTable data
    /// </summary>
    public List<List<object?>> Values { get; set; } = new();
    
    /// <summary>
    /// Column headers
    /// </summary>
    public List<string> ColumnHeaders { get; set; } = new();
    
    /// <summary>
    /// Row headers
    /// </summary>
    public List<string> RowHeaders { get; set; } = new();
    
    /// <summary>
    /// Number of data rows
    /// </summary>
    public int DataRowCount { get; set; }
    
    /// <summary>
    /// Number of data columns
    /// </summary>
    public int DataColumnCount { get; set; }
    
    /// <summary>
    /// Grand totals
    /// </summary>
    public Dictionary<string, object?> GrandTotals { get; set; } = new();
}

/// <summary>
/// Result for field filter operations
/// </summary>
public class PivotFieldFilterResult : ResultBase
{
    /// <summary>
    /// Name of the field
    /// </summary>
    public string FieldName { get; set; } = string.Empty;
    
    /// <summary>
    /// Selected items
    /// </summary>
    public List<string> SelectedItems { get; set; } = new();
    
    /// <summary>
    /// Available items
    /// </summary>
    public List<string> AvailableItems { get; set; } = new();
    
    /// <summary>
    /// Number of visible rows after filter
    /// </summary>
    public int VisibleRowCount { get; set; }
    
    /// <summary>
    /// Total rows before filter
    /// </summary>
    public int TotalRowCount { get; set; }
    
    /// <summary>
    /// Whether all items are shown
    /// </summary>
    public bool ShowAll { get; set; }
}
