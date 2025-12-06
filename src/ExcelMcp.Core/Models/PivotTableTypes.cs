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
/// Excel COM constants for CubeField.CubeFieldType property.
/// Used to distinguish between different types of OLAP CubeFields.
/// </summary>
public static class XlCubeFieldType
{
    /// <summary>
    /// Hierarchy field (table column/dimension)
    /// </summary>
    public const int xlHierarchy = 1;

    /// <summary>
    /// Measure field (DAX measure or implicit measure)
    /// </summary>
    public const int xlMeasure = 2;

    /// <summary>
    /// Set field
    /// </summary>
    public const int xlSet = 3;
}

/// <summary>
/// Excel COM constants for sort order
/// </summary>
public static class XlSortOrder
{
    /// <summary>
    /// Sort ascending
    /// </summary>
    public const int xlAscending = 1;

    /// <summary>
    /// Sort descending
    /// </summary>
    public const int xlDescending = 2;
}

/// <summary>
/// Excel PivotField data type constants
/// </summary>
public static class XlPivotFieldDataType
{
    /// <summary>
    /// Date field type
    /// </summary>
    public const int xlDate = 2;

    /// <summary>
    /// Number field type
    /// </summary>
    public const int xlNumber = -4145;

    /// <summary>
    /// Text field type
    /// </summary>
    public const int xlText = -4158;
}

/// <summary>
/// Excel time unit constants for date grouping
/// </summary>
public static class XlTimeUnit
{
    /// <summary>
    /// Days grouping
    /// </summary>
    public const int xlDays = 4;

    /// <summary>
    /// Months grouping
    /// </summary>
    public const int xlMonths = 5;

    /// <summary>
    /// Quarters grouping
    /// </summary>
    public const int xlQuarters = 6;

    /// <summary>
    /// Years grouping
    /// </summary>
    public const int xlYears = 7;
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
    /// All available fields from source data that can be added to the PivotTable
    /// </summary>
    public List<string> AvailableFields { get; set; } = [];
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
    public List<PivotTableInfo> PivotTables { get; set; } = [];
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
    public List<PivotFieldInfo> Fields { get; set; } = [];
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

    /// <summary>
    /// Formula for calculated fields (e.g., "=Revenue-Cost")
    /// </summary>
    public string? Formula { get; set; }
}

/// <summary>
/// Result for field listing operations
/// </summary>
public class PivotFieldListResult : ResultBase
{
    /// <summary>
    /// List of all fields in the PivotTable
    /// </summary>
    public List<PivotFieldInfo> Fields { get; set; } = [];
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
    public List<string> AvailableValues { get; set; } = [];

    /// <summary>
    /// Sample value for verification
    /// </summary>
    public object? SampleValue { get; set; }

    /// <summary>
    /// Data type of the field
    /// </summary>
    public string DataType { get; set; } = string.Empty;

    /// <summary>
    /// Formula for calculated fields (e.g., "=Revenue-Cost")
    /// </summary>
    public string? Formula { get; set; }

    /// <summary>
    /// Workflow hint describing what happened and suggested next steps
    /// </summary>
    public string? WorkflowHint { get; set; }
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
    public List<string> NewFields { get; set; } = [];

    /// <summary>
    /// Fields removed from source
    /// </summary>
    public List<string> RemovedFields { get; set; } = [];
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
    public List<List<object?>> Values { get; set; } = [];

    /// <summary>
    /// Column headers
    /// </summary>
    public List<string> ColumnHeaders { get; set; } = [];

    /// <summary>
    /// Row headers
    /// </summary>
    public List<string> RowHeaders { get; set; } = [];

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
    public Dictionary<string, object?> GrandTotals { get; set; } = [];
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
    public List<string> SelectedItems { get; set; } = [];

    /// <summary>
    /// Available items
    /// </summary>
    public List<string> AvailableItems { get; set; } = [];

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

/// <summary>
/// Information about a calculated field in a regular PivotTable
/// </summary>
public class CalculatedFieldInfo
{
    /// <summary>
    /// Name of the calculated field
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Formula for the calculated field (e.g., "=Revenue-Cost")
    /// </summary>
    public string Formula { get; set; } = string.Empty;

    /// <summary>
    /// Source name of the field
    /// </summary>
    public string? SourceName { get; set; }
}

/// <summary>
/// Result for listing calculated fields
/// </summary>
public class CalculatedFieldListResult : ResultBase
{
    /// <summary>
    /// List of calculated fields in the PivotTable
    /// </summary>
    public List<CalculatedFieldInfo> CalculatedFields { get; set; } = [];
}

/// <summary>
/// Excel COM constants for calculated member types
/// </summary>
public static class XlCalculatedMemberType
{
    /// <summary>
    /// Calculated member (custom MDX formula member)
    /// </summary>
    public const int xlCalculatedMember = 0;

    /// <summary>
    /// Calculated set (named set of members)
    /// </summary>
    public const int xlCalculatedSet = 1;

    /// <summary>
    /// Calculated measure (DAX-like measure for Data Model)
    /// </summary>
    public const int xlCalculatedMeasure = 2;
}

/// <summary>
/// Type of calculated member
/// </summary>
public enum CalculatedMemberType
{
    /// <summary>
    /// Calculated member (custom MDX formula member)
    /// </summary>
    Member = 0,

    /// <summary>
    /// Calculated set (named set of members)
    /// </summary>
    Set = 1,

    /// <summary>
    /// Calculated measure (DAX-like measure for Data Model)
    /// </summary>
    Measure = 2
}

/// <summary>
/// Information about a calculated member in a PivotTable
/// </summary>
public class CalculatedMemberInfo
{
    /// <summary>
    /// Name of the calculated member
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// MDX or DAX formula
    /// </summary>
    public string Formula { get; set; } = string.Empty;

    /// <summary>
    /// Type of calculated member (Member, Set, or Measure)
    /// </summary>
    public CalculatedMemberType Type { get; set; }

    /// <summary>
    /// Solve order for calculation precedence
    /// </summary>
    public int SolveOrder { get; set; }

    /// <summary>
    /// Whether the calculated member is valid
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// Display folder path (for measures)
    /// </summary>
    public string? DisplayFolder { get; set; }

    /// <summary>
    /// Number format code
    /// </summary>
    public string? NumberFormat { get; set; }
}

/// <summary>
/// Result for listing calculated members
/// </summary>
public class CalculatedMemberListResult : ResultBase
{
    /// <summary>
    /// List of calculated members in the PivotTable
    /// </summary>
    public List<CalculatedMemberInfo> CalculatedMembers { get; set; } = [];
}

/// <summary>
/// Result for calculated member operations
/// </summary>
public class CalculatedMemberResult : ResultBase
{
    /// <summary>
    /// Name of the calculated member
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// MDX or DAX formula
    /// </summary>
    public string Formula { get; set; } = string.Empty;

    /// <summary>
    /// Type of calculated member
    /// </summary>
    public CalculatedMemberType Type { get; set; }

    /// <summary>
    /// Solve order for calculation precedence
    /// </summary>
    public int SolveOrder { get; set; }

    /// <summary>
    /// Whether the calculated member is valid
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// Display folder path (for measures)
    /// </summary>
    public string? DisplayFolder { get; set; }

    /// <summary>
    /// Number format code
    /// </summary>
    public string? NumberFormat { get; set; }

    /// <summary>
    /// Workflow hint describing what happened and suggested next steps
    /// </summary>
    public string? WorkflowHint { get; set; }
}
