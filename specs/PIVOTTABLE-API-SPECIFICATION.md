# Excel PivotTable API Specification

> **Comprehensive specification for Excel PivotTable operations - creating, managing, and analyzing data with pivot functionality**

## Executive Summary

This specification defines a **PivotTable API** for ExcelMcp that provides complete PivotTable lifecycle management, field configuration, and data analysis capabilities through Excel COM automation.

### Key Design Decisions

1. **COM-Backed Only** - Every operation uses native Excel COM PivotTable objects
2. **Data Source Agnostic** - Support Excel ranges, tables, external connections, and Data Model
3. **Field-Centric Design** - Operations organized around PivotTable fields and areas
4. **Cache Management** - Proper handling of PivotCache for performance and data refresh
5. **Layout Preservation** - Maintain existing PivotTable structure during modifications

### Goals

1. **Complete Lifecycle** - Create, configure, refresh, and delete PivotTables
2. **Field Management** - Add/remove/configure fields in all areas (Rows, Columns, Values, Filters)
3. **Data Analysis** - Sorting, filtering, grouping, and drill-down capabilities
4. **Performance** - Efficient cache management and bulk operations
5. **Power User Features** - Advanced formatting, calculated fields, and slicers

---

## Excel PivotTable Architecture

### What is a PivotTable?

Excel PivotTables are **interactive data summarization tools** that provide:
- Dynamic data aggregation and cross-tabulation
- Drag-and-drop field configuration
- Multiple aggregation functions (Sum, Count, Average, etc.)
- Hierarchical grouping and drilling
- Interactive filtering and slicing
- Calculated fields and items
- Professional formatting and styling

### Excel COM Object Model

#### Core Objects
```csharp
// PivotTable object hierarchy
dynamic worksheet = workbook.Worksheets.Item("Sheet1");
dynamic pivotTables = worksheet.PivotTables;
dynamic pivotTable = pivotTables.Item("PivotTable1");

// PivotCache (data source)
dynamic pivotCache = pivotTable.PivotCache;

// PivotFields (columns from source data)
dynamic pivotFields = pivotTable.PivotFields;
dynamic field = pivotFields.Item("Sales");

// Field Areas
dynamic rowFields = pivotTable.RowFields;      // Row area
dynamic columnFields = pivotTable.ColumnFields; // Column area
dynamic dataFields = pivotTable.DataFields;     // Values area
dynamic pageFields = pivotTable.PageFields;     // Filter area
```

### PivotTable Creation Workflow

```csharp
// VALIDATED COM API PATTERNS (from Excel VBA documentation)

// 1. Create PivotCache from data source
dynamic pivotCaches = workbook.PivotCaches();
dynamic pivotCache = pivotCaches.Create(
    SourceType: 1,                    // xlDatabase = 1
    SourceData: "Sheet1!A1:F100"     // Data range with headers
);

// 2. Create PivotTable from cache  
dynamic destinationSheet = workbook.Worksheets.Item("Sheet2");
dynamic pivotTable = pivotCache.CreatePivotTable(
    TableDestination: destinationSheet.Range["A1"],  // Range object, not string
    TableName: "SalesPivot"
);

// 3. Configure fields with proper constants
dynamic salesField = pivotTable.PivotFields.Item("Sales");
salesField.Orientation = 4;          // xlDataField = 4
salesField.Function = -4157;         // xlSum = -4157

dynamic regionField = pivotTable.PivotFields.Item("Region");
regionField.Orientation = 1;         // xlRowField = 1

// 4. CRITICAL: Refresh to materialize layout
pivotTable.RefreshTable();
```

---

## Proposed PivotTable API Design

### Design Principles

1. **COM-Backed Only**: Every method uses native Excel COM PivotTable operations
2. **Cache-Aware**: Proper PivotCache management for performance and data integrity
3. **Field-Centric**: Operations organized around PivotTable field management
4. **Incremental Configuration**: Build PivotTables step-by-step with validation
5. **Error Recovery**: Handle invalid field configurations gracefully

## LLM-Optimized API Design

### Design Principles for AI Agents

1. **COM-Backed with Validated Constants**: Every method uses correct Excel COM constants and error handling
2. **Meaningful Result Validation**: Integration tests verify actual PivotTable structure, not just success status
3. **Incremental Build Pattern**: LLMs can build PivotTables step-by-step with immediate feedback
4. **Error Recovery**: Clear error messages with actionable guidance for field configuration issues
5. **Data Source Transparency**: Auto-detect and validate data source types (range, table, Data Model)

### LLM Usage Patterns

**As an LLM, I need to:**

1. **Create PivotTable from existing data** - Auto-detect data boundaries and headers
2. **Add fields incrementally** - Get immediate feedback on field placement and validation
3. **Verify layout changes** - Read back PivotTable structure after each modification
4. **Handle configuration errors** - Graceful recovery when fields don't exist or have wrong types
5. **Access result data** - Read PivotTable output for analysis and reporting

### Phase 1: LLM-First Core Operations (MVP)

#### IPivotTableCommands Interface (Validated COM Patterns)

```csharp
public interface IPivotTableCommands
{
    // === LIFECYCLE OPERATIONS ===
    
    /// <summary>
    /// Lists all PivotTables in workbook with structure details
    /// Returns: Name, sheet, range, source data, field counts, last refresh
    /// </summary>
    Task<PivotTableListResult> ListAsync(IExcelBatch batch);
    
    /// <summary>
    /// Gets complete PivotTable configuration and current layout
    /// Returns: All fields with positions, aggregation functions, filter states
    /// </summary>
    Task<PivotTableInfoResult> GetInfoAsync(IExcelBatch batch, string pivotTableName);
    
    /// <summary>
    /// Creates PivotTable from Excel range with auto-detection of headers
    /// Validates: Source range exists, has headers, contains data
    /// Returns: Created PivotTable name and initial field list
    /// </summary>
    Task<PivotTableCreateResult> CreateFromRangeAsync(IExcelBatch batch, 
        string sourceSheet, string sourceRange, 
        string destinationSheet, string destinationCell, 
        string pivotTableName);
    
    /// <summary>
    /// Creates PivotTable from Excel Table (ListObject)
    /// Validates: Table exists, has data rows
    /// Returns: Created PivotTable name and available fields
    /// </summary>
    Task<PivotTableCreateResult> CreateFromTableAsync(IExcelBatch batch, 
        string tableName, 
        string destinationSheet, string destinationCell, 
        string pivotTableName);
    
    /// <summary>
    /// Deletes PivotTable completely
    /// Validates: PivotTable exists before deletion
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string pivotTableName);
    
    /// <summary>
    /// Refreshes PivotTable data from source and returns updated info
    /// Returns: Refresh timestamp, record count, any structural changes
    /// </summary>
    Task<PivotTableRefreshResult> RefreshAsync(IExcelBatch batch, string pivotTableName);
    
    // === FIELD MANAGEMENT (WITH IMMEDIATE VALIDATION) ===
    
    /// <summary>
    /// Lists all available fields and their current placement
    /// Returns: Field names, data types, current areas, aggregation functions
    /// </summary>
    Task<PivotFieldListResult> ListFieldsAsync(IExcelBatch batch, string pivotTableName);
    
    /// <summary>
    /// Adds field to Row area with position validation
    /// Validates: Field exists, not already in another area
    /// Returns: Updated field layout with new position
    /// </summary>
    Task<PivotFieldResult> AddRowFieldAsync(IExcelBatch batch, string pivotTableName, 
        string fieldName, int? position = null);
    
    /// <summary>
    /// Adds field to Column area with position validation
    /// Validates: Field exists, not already in another area
    /// Returns: Updated field layout
    /// </summary>
    Task<PivotFieldResult> AddColumnFieldAsync(IExcelBatch batch, string pivotTableName, 
        string fieldName, int? position = null);
    
    /// <summary>
    /// Adds field to Values area with aggregation function
    /// Validates: Field exists, aggregation function is appropriate for data type
    /// Returns: Field configuration with applied function and custom name
    /// </summary>
    Task<PivotFieldResult> AddValueFieldAsync(IExcelBatch batch, string pivotTableName, 
        string fieldName, AggregationFunction function = AggregationFunction.Sum, 
        string? customName = null);
    
    /// <summary>
    /// Adds field to Filter area (Page field)
    /// Validates: Field exists, returns available filter values
    /// Returns: Field configuration with available filter items
    /// </summary>
    Task<PivotFieldResult> AddFilterFieldAsync(IExcelBatch batch, string pivotTableName, 
        string fieldName);
    
    /// <summary>
    /// Removes field from any area
    /// Validates: Field is currently placed in PivotTable
    /// Returns: Updated layout after removal
    /// </summary>
    Task<PivotFieldResult> RemoveFieldAsync(IExcelBatch batch, string pivotTableName, 
        string fieldName);
    
    /// <summary>
    /// Moves field between areas with validation
    /// Validates: Valid source/target areas, position constraints
    /// Returns: New field configuration and layout
    /// </summary>
    Task<PivotFieldResult> MoveFieldAsync(IExcelBatch batch, string pivotTableName, 
        string fieldName, PivotFieldArea fromArea, PivotFieldArea toArea, 
        int? position = null);
    
    // === FIELD CONFIGURATION (WITH RESULT VERIFICATION) ===
    
    /// <summary>
    /// Sets aggregation function for value field
    /// Validates: Field is in Values area, function is valid for data type
    /// Returns: Applied function and sample calculation result
    /// </summary>
    Task<PivotFieldResult> SetFieldFunctionAsync(IExcelBatch batch, string pivotTableName, 
        string fieldName, AggregationFunction function);
    
    /// <summary>
    /// Sets custom name for field in any area
    /// Validates: Name doesn't conflict with existing fields
    /// Returns: Applied name and field reference
    /// </summary>
    Task<PivotFieldResult> SetFieldNameAsync(IExcelBatch batch, string pivotTableName, 
        string fieldName, string customName);
    
    /// <summary>
    /// Sets number format for value field
    /// Validates: Field is in Values area, format string is valid
    /// Returns: Applied format with sample formatted value
    /// </summary>
    Task<PivotFieldResult> SetFieldFormatAsync(IExcelBatch batch, string pivotTableName, 
        string fieldName, string numberFormat);
    
    // === ANALYSIS OPERATIONS (WITH DATA VALIDATION) ===
    
    /// <summary>
    /// Gets current PivotTable data as 2D array for LLM analysis
    /// Returns: Values with headers, row/column labels, formatted numbers
    /// </summary>
    Task<PivotTableDataResult> GetDataAsync(IExcelBatch batch, string pivotTableName);
    
    /// <summary>
    /// Sets filter for field with validation of filter items
    /// Validates: Field supports filtering, selected values exist
    /// Returns: Applied filter state and affected row count
    /// </summary>
    Task<PivotFieldFilterResult> SetFieldFilterAsync(IExcelBatch batch, string pivotTableName, 
        string fieldName, List<string> selectedValues);
    
    /// <summary>
    /// Sorts field with immediate layout update
    /// Validates: Field can be sorted, returns new sort order
    /// Returns: Applied sort configuration and preview of changes
    /// </summary>
    Task<PivotFieldResult> SortFieldAsync(IExcelBatch batch, string pivotTableName, 
        string fieldName, SortDirection direction = SortDirection.Ascending);
}

// === LLM-FOCUSED RESULT TYPES ===

public class PivotTableCreateResult : OperationResult
{
    public string PivotTableName { get; set; } = string.Empty;
    public string SheetName { get; set; } = string.Empty;
    public string Range { get; set; } = string.Empty;
    public string SourceData { get; set; } = string.Empty;
    public int SourceRowCount { get; set; }
    public List<string> AvailableFields { get; set; } = new();
    public List<string> NumericFields { get; set; } = new();  // Suggested for Values area
    public List<string> TextFields { get; set; } = new();     // Suggested for Rows/Columns/Filters
    public List<string> DateFields { get; set; } = new();     // Suggested for grouping
}

public class PivotTableRefreshResult : OperationResult
{
    public string PivotTableName { get; set; } = string.Empty;
    public DateTime RefreshTime { get; set; }
    public int SourceRecordCount { get; set; }
    public int PreviousRecordCount { get; set; }
    public bool StructureChanged { get; set; }
    public List<string> NewFields { get; set; } = new();      // Fields added to source
    public List<string> RemovedFields { get; set; } = new();  // Fields removed from source
}

public class PivotFieldResult : OperationResult
{
    public string FieldName { get; set; } = string.Empty;
    public string CustomName { get; set; } = string.Empty;
    public PivotFieldArea Area { get; set; }
    public int Position { get; set; }
    public AggregationFunction? Function { get; set; }
    public string? NumberFormat { get; set; }
    public List<string> AvailableValues { get; set; } = new();  // For filtering
    public object? SampleValue { get; set; }                   // Example of formatted output
    public string DataType { get; set; } = string.Empty;       // Text, Number, Date, Boolean
}

public class PivotTableDataResult : OperationResult
{
    public string PivotTableName { get; set; } = string.Empty;
    public List<List<object?>> Values { get; set; } = new();   // 2D array of PivotTable data
    public List<string> ColumnHeaders { get; set; } = new();   // Column field values
    public List<string> RowHeaders { get; set; } = new();      // Row field values  
    public int DataRowCount { get; set; }
    public int DataColumnCount { get; set; }
    public Dictionary<string, object?> GrandTotals { get; set; } = new();  // Summary values
}

public class PivotFieldFilterResult : OperationResult
{
    public string FieldName { get; set; } = string.Empty;
    public List<string> SelectedItems { get; set; } = new();
    public List<string> AvailableItems { get; set; } = new();
    public int VisibleRowCount { get; set; }        // Rows shown after filter
    public int TotalRowCount { get; set; }          // Total rows before filter
    public bool ShowAll { get; set; }
}
```

### Phase 2: Advanced Operations

```csharp
public interface IPivotTableCommands
{
    // === GROUPING OPERATIONS ===
    
    /// <summary>
    /// Groups date field by period (years, quarters, months, days)
    /// </summary>
    Task<OperationResult> GroupDateFieldAsync(IExcelBatch batch, string pivotTableName, string fieldName, 
        DateGrouping groupBy, DateTime? startDate = null, DateTime? endDate = null);
    
    /// <summary>
    /// Groups numeric field by ranges
    /// </summary>
    Task<OperationResult> GroupNumericFieldAsync(IExcelBatch batch, string pivotTableName, string fieldName, 
        double? startValue, double? endValue, double interval);
    
    /// <summary>
    /// Creates custom grouping for text field
    /// </summary>
    Task<OperationResult> CreateCustomGroupAsync(IExcelBatch batch, string pivotTableName, string fieldName, 
        string groupName, List<string> items);
    
    /// <summary>
    /// Ungroups a field
    /// </summary>
    Task<OperationResult> UngroupFieldAsync(IExcelBatch batch, string pivotTableName, string fieldName);
    
    // === CALCULATED FIELDS ===
    
    /// <summary>
    /// Creates a calculated field with custom formula
    /// </summary>
    Task<OperationResult> CreateCalculatedFieldAsync(IExcelBatch batch, string pivotTableName, 
        string fieldName, string formula);
    
    /// <summary>
    /// Updates calculated field formula
    /// </summary>
    Task<OperationResult> UpdateCalculatedFieldAsync(IExcelBatch batch, string pivotTableName, 
        string fieldName, string formula);
    
    /// <summary>
    /// Deletes calculated field
    /// </summary>
    Task<OperationResult> DeleteCalculatedFieldAsync(IExcelBatch batch, string pivotTableName, string fieldName);
    
    /// <summary>
    /// Lists all calculated fields
    /// </summary>
    Task<CalculatedFieldListResult> ListCalculatedFieldsAsync(IExcelBatch batch, string pivotTableName);
    
    // === DRILL DOWN ===
    
    /// <summary>
    /// Gets drill-down data for a specific cell in PivotTable
    /// </summary>
    Task<PivotDrillDownResult> DrillDownAsync(IExcelBatch batch, string pivotTableName, 
        string targetSheet, string cellAddress);
    
    // === SLICER INTEGRATION ===
    
    /// <summary>
    /// Creates a slicer for a PivotTable field
    /// </summary>
    Task<OperationResult> CreateSlicerAsync(IExcelBatch batch, string pivotTableName, string fieldName, 
        string slicerName, string destinationSheet, string position);
    
    /// <summary>
    /// Lists all slicers connected to PivotTable
    /// </summary>
    Task<SlicerListResult> ListSlicersAsync(IExcelBatch batch, string pivotTableName);
    
    /// <summary>
    /// Sets slicer selection
    /// </summary>
    Task<OperationResult> SetSlicerSelectionAsync(IExcelBatch batch, string slicerName, List<string> selectedItems);
    
    // === DATA SOURCE MANAGEMENT ===
    
    /// <summary>
    /// Changes PivotTable data source
    /// </summary>
    Task<OperationResult> ChangeDataSourceAsync(IExcelBatch batch, string pivotTableName, 
        string newSourceSheet, string newSourceRange);
    
    /// <summary>
    /// Gets PivotCache information
    /// </summary>
    Task<PivotCacheInfoResult> GetCacheInfoAsync(IExcelBatch batch, string pivotTableName);
}

// === ADDITIONAL SUPPORTING TYPES ===

public enum DateGrouping
{
    Years,
    Quarters,
    Months,
    Days,
    Hours,
    Minutes,
    Seconds
}

public class CalculatedFieldInfo
{
    public string Name { get; set; } = string.Empty;
    public string Formula { get; set; } = string.Empty;
}

public class CalculatedFieldListResult : OperationResult
{
    public List<CalculatedFieldInfo> CalculatedFields { get; set; } = new();
}

public class PivotDrillDownResult : OperationResult
{
    public string DrillDownSheet { get; set; } = string.Empty;
    public string DrillDownRange { get; set; } = string.Empty;
    public int RecordCount { get; set; }
    public List<string> ColumnHeaders { get; set; } = new();
}

public class SlicerInfo
{
    public string Name { get; set; } = string.Empty;
    public string FieldName { get; set; } = string.Empty;
    public string SheetName { get; set; } = string.Empty;
    public List<string> SelectedItems { get; set; } = new();
}

public class SlicerListResult : OperationResult
{
    public List<SlicerInfo> Slicers { get; set; } = new();
}

public class PivotCacheInfo
{
    public string SourceData { get; set; } = string.Empty;
    public DateTime LastRefresh { get; set; }
    public int RecordCount { get; set; }
    public List<string> FieldNames { get; set; } = new();
}

public class PivotCacheInfoResult : OperationResult
{
    public PivotCacheInfo CacheInfo { get; set; } = new();
}
```

---

## Excel COM Implementation Details (VALIDATED)

### PivotTable Creation Patterns (COM API Verified)

#### From Excel Range (Most Common)
```csharp
// STEP 1: Validate source data
dynamic sourceSheet = workbook.Worksheets.Item(sourceSheetName);
dynamic sourceRange = sourceSheet.Range[sourceRangeAddress];

// Validation checks
if (sourceRange.Rows.Count < 2)
    throw new McpException($"Source range must contain headers and at least one data row. Found {sourceRange.Rows.Count} rows");

// Check for headers in first row
object[,] headerRow = sourceRange.Rows[1].Value2;
var headers = new List<string>();
for (int col = 1; col <= headerRow.GetLength(1); col++)
{
    var header = headerRow[1, col]?.ToString();
    if (string.IsNullOrWhiteSpace(header))
        throw new McpException($"Missing header in column {col}. All columns must have headers.");
    headers.Add(header);
}

// STEP 2: Create PivotCache with error handling
dynamic pivotCaches = workbook.PivotCaches();
dynamic pivotCache;
try 
{
    pivotCache = pivotCaches.Create(
        SourceType: 1,  // xlDatabase = 1 (VALIDATED CONSTANT)
        SourceData: $"{sourceSheetName}!{sourceRangeAddress}"
    );
}
catch (Exception ex)
{
    throw new McpException($"Failed to create PivotCache from {sourceSheetName}!{sourceRangeAddress}: {ex.Message}");
}

// STEP 3: Create PivotTable with destination validation
dynamic destinationSheet = workbook.Worksheets.Item(destinationSheetName);
dynamic destinationRange = destinationSheet.Range[destinationCell];

dynamic pivotTable;
try 
{
    pivotTable = pivotCache.CreatePivotTable(
        TableDestination: destinationRange,  // Must be Range object, not string
        TableName: pivotTableName
    );
}
catch (Exception ex)
{
    ComUtilities.Release(ref pivotCache);
    throw new McpException($"Failed to create PivotTable '{pivotTableName}' at {destinationSheetName}!{destinationCell}: {ex.Message}");
}

// STEP 4: CRITICAL - Refresh to materialize layout
try 
{
    pivotTable.RefreshTable();
}
catch (Exception ex)
{
    throw new McpException($"Failed to refresh PivotTable '{pivotTableName}': {ex.Message}");
}

// Return creation result with validation
return new PivotTableCreateResult
{
    Success = true,
    PivotTableName = pivotTableName,
    SheetName = destinationSheetName,
    Range = pivotTable.TableRange2.Address,
    SourceData = sourceRangeAddress,
    SourceRowCount = sourceRange.Rows.Count - 1, // Exclude headers
    AvailableFields = headers,
    NumericFields = DetectNumericFields(sourceRange, headers),
    TextFields = DetectTextFields(sourceRange, headers),
    DateFields = DetectDateFields(sourceRange, headers)
};
```

#### Field Management with COM Constants (VALIDATED)

```csharp
// EXCEL COM CONSTANTS (from Excel VBA documentation)
public static class XlPivotFieldOrientation 
{
    public const int xlHidden = 0;      // Field not displayed
    public const int xlRowField = 1;    // Row area
    public const int xlColumnField = 2; // Column area  
    public const int xlPageField = 3;   // Filter area (Page field)
    public const int xlDataField = 4;   // Values area
}

public static class XlConsolidationFunction
{
    public const int xlSum = -4157;
    public const int xlCount = -4112;
    public const int xlAverage = -4106;
    public const int xlMax = -4136;
    public const int xlMin = -4139;
    public const int xlProduct = -4149;
    public const int xlCountNums = -4113;
    public const int xlStdDev = -4155;
    public const int xlStdDevP = -4156;
    public const int xlVar = -4164;
    public const int xlVarP = -4165;
}

// Adding field to Row area with validation
public async Task<PivotFieldResult> AddRowFieldAsync(IExcelBatch batch, 
    string pivotTableName, string fieldName, int? position = null)
{
    return await batch.ExecuteAsync(async (ctx, ct) =>
    {
        // Find PivotTable
        dynamic pivotTable = FindPivotTable(ctx.Book, pivotTableName);
        
        // Validate field exists
        dynamic field;
        try 
        {
            field = pivotTable.PivotFields.Item(fieldName);
        }
        catch (Exception)
        {
            throw new McpException($"Field '{fieldName}' not found in PivotTable '{pivotTableName}'. Available fields: {string.Join(", ", GetFieldNames(pivotTable))}");
        }
        
        // Check if field is already placed
        int currentOrientation = field.Orientation;
        if (currentOrientation != XlPivotFieldOrientation.xlHidden)
        {
            throw new McpException($"Field '{fieldName}' is already placed in {GetAreaName(currentOrientation)} area. Remove it first or use MoveField.");
        }
        
        // Add to Row area
        try 
        {
            field.Orientation = XlPivotFieldOrientation.xlRowField;
            if (position.HasValue)
            {
                field.Position = position.Value;
            }
        }
        catch (Exception ex)
        {
            throw new McpException($"Failed to add field '{fieldName}' to Row area: {ex.Message}");
        }
        
        // Refresh and validate placement
        pivotTable.RefreshTable();
        
        // Verify field was added successfully
        if (field.Orientation != XlPivotFieldOrientation.xlRowField)
        {
            throw new McpException($"Field '{fieldName}' was not successfully added to Row area. Current orientation: {GetAreaName(field.Orientation)}");
        }
        
        // Return detailed result
        return new PivotFieldResult
        {
            Success = true,
            FieldName = fieldName,
            CustomName = field.Caption?.ToString() ?? fieldName,
            Area = PivotFieldArea.Row,
            Position = field.Position,
            DataType = DetectFieldDataType(field),
            AvailableValues = GetFieldUniqueValues(field),
            SampleValue = GetFieldSampleValue(field)
        };
    });
}

// Value field with aggregation function validation
public async Task<PivotFieldResult> AddValueFieldAsync(IExcelBatch batch, 
    string pivotTableName, string fieldName, 
    AggregationFunction function = AggregationFunction.Sum, 
    string? customName = null)
{
    return await batch.ExecuteAsync(async (ctx, ct) =>
    {
        dynamic pivotTable = FindPivotTable(ctx.Book, pivotTableName);
        dynamic field = pivotTable.PivotFields.Item(fieldName);
        
        // Validate aggregation function for field data type
        string dataType = DetectFieldDataType(field);
        if (!IsValidAggregationForDataType(function, dataType))
        {
            var validFunctions = GetValidAggregationsForDataType(dataType);
            throw new McpException($"Aggregation function '{function}' is not valid for {dataType} field '{fieldName}'. Valid functions: {string.Join(", ", validFunctions)}");
        }
        
        // Add to Values area
        field.Orientation = XlPivotFieldOrientation.xlDataField;
        
        // Set aggregation function with COM constant
        int comFunction = GetComAggregationFunction(function);
        field.Function = comFunction;
        
        // Set custom name if provided
        if (!string.IsNullOrEmpty(customName))
        {
            field.Caption = customName;
        }
        
        // Refresh and validate
        pivotTable.RefreshTable();
        
        // Get sample calculated value for verification
        object? sampleValue = GetValueFieldSample(pivotTable, fieldName, function);
        
        return new PivotFieldResult
        {
            Success = true,
            FieldName = fieldName,
            CustomName = field.Caption?.ToString() ?? fieldName,
            Area = PivotFieldArea.Value,
            Function = function,
            DataType = dataType,
            SampleValue = sampleValue
        };
    });
}
```

### Data Type Detection and Validation

```csharp
// Field data type detection for validation
private string DetectFieldDataType(dynamic field)
{
    try 
    {
        // Get sample values from field
        dynamic pivotItems = field.PivotItems;
        var sampleValues = new List<object>();
        
        int sampleCount = Math.Min(10, pivotItems.Count);
        for (int i = 1; i <= sampleCount; i++)
        {
            var value = pivotItems.Item(i).Value;
            if (value != null)
                sampleValues.Add(value);
        }
        
        // Analyze sample values
        if (sampleValues.All(v => DateTime.TryParse(v.ToString(), out _)))
            return "Date";
        if (sampleValues.All(v => double.TryParse(v.ToString(), out _)))
            return "Number";
        if (sampleValues.All(v => bool.TryParse(v.ToString(), out _)))
            return "Boolean";
        
        return "Text";
    }
    catch 
    {
        return "Unknown";
    }
}

// Aggregation function validation
private bool IsValidAggregationForDataType(AggregationFunction function, string dataType)
{
    return dataType switch
    {
        "Number" => true, // All functions valid for numbers
        "Date" => function == AggregationFunction.Count || function == AggregationFunction.CountNumbers ||
                  function == AggregationFunction.Max || function == AggregationFunction.Min,
        "Text" => function == AggregationFunction.Count,
        "Boolean" => function == AggregationFunction.Count || function == AggregationFunction.Sum,
        _ => function == AggregationFunction.Count
    };
}

// COM constant mapping with validation
private int GetComAggregationFunction(AggregationFunction function)
{
    return function switch
    {
        AggregationFunction.Sum => XlConsolidationFunction.xlSum,
        AggregationFunction.Count => XlConsolidationFunction.xlCount,
        AggregationFunction.Average => XlConsolidationFunction.xlAverage,
        AggregationFunction.Max => XlConsolidationFunction.xlMax,
        AggregationFunction.Min => XlConsolidationFunction.xlMin,
        AggregationFunction.Product => XlConsolidationFunction.xlProduct,
        AggregationFunction.CountNumbers => XlConsolidationFunction.xlCountNums,
        AggregationFunction.StdDev => XlConsolidationFunction.xlStdDev,
        AggregationFunction.StdDevP => XlConsolidationFunction.xlStdDevP,
        AggregationFunction.Var => XlConsolidationFunction.xlVar,
        AggregationFunction.VarP => XlConsolidationFunction.xlVarP,
        _ => throw new McpException($"Unsupported aggregation function: {function}")
    };
}
```

### Error Recovery and Cleanup Patterns

```csharp
// Robust PivotTable creation with cleanup
public async Task<PivotTableCreateResult> CreateFromRangeAsync(IExcelBatch batch, 
    string sourceSheet, string sourceRange, 
    string destinationSheet, string destinationCell, 
    string pivotTableName)
{
    dynamic? pivotCache = null;
    dynamic? pivotTable = null;
    
    try 
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            // Validation and creation logic here...
            
            return result;
        });
    }
    catch (Exception ex)
    {
        // Cleanup on failure
        if (pivotTable != null)
        {
            try { pivotTable.Delete(); } catch { }
            ComUtilities.Release(ref pivotTable);
        }
        
        if (pivotCache != null)
        {
            try { pivotCache.Delete(); } catch { }
            ComUtilities.Release(ref pivotCache);
        }
        
        throw new McpException($"PivotTable creation failed: {ex.Message}", ex);
    }
}
```

## Integration Testing Strategy (Result Validation)

### Test Philosophy: Verify Actual Results, Not Just Success Status

**❌ WRONG - Testing Only Success Status:**
```csharp
[Fact]
public async Task CreatePivotTable_ValidData_ReturnsSuccess()
{
    var result = await _commands.CreateFromRangeAsync(batch, "Data", "A1:D100", "Summary", "A1", "TestPivot");
    Assert.True(result.Success); // ❌ Only tests success flag
}
```

**✅ CORRECT - Testing Actual Excel State:**
```csharp
[Fact]
public async Task CreatePivotTable_ValidData_CreatesCorrectPivotStructure()
{
    // Arrange - Create test data with known structure
    await SetupTestData(batch, "Data", new[]
    {
        new[] { "Region", "Product", "Sales", "Date" },
        new[] { "North", "A", "100", "2024-01-01" },
        new[] { "South", "B", "200", "2024-01-02" },
        new[] { "North", "A", "150", "2024-01-03" }
    });

    // Act
    var result = await _commands.CreateFromRangeAsync(batch, "Data", "A1:D4", "Summary", "A1", "TestPivot");

    // Assert - Verify Success AND Structure
    Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
    Assert.Equal("TestPivot", result.PivotTableName);
    Assert.Equal(4, result.AvailableFields.Count);
    Assert.Contains("Region", result.AvailableFields);
    Assert.Contains("Product", result.AvailableFields);
    Assert.Contains("Sales", result.NumericFields);  // Should detect as numeric
    Assert.Contains("Date", result.DateFields);      // Should detect as date
    Assert.Contains("Region", result.TextFields);    // Should detect as text
    
    // Verify actual Excel COM object exists and is configured correctly
    await VerifyPivotTableExists(batch, "TestPivot", "Summary");
}

private async Task VerifyPivotTableExists(IExcelBatch batch, string pivotTableName, string sheetName)
{
    await batch.ExecuteAsync(async (ctx, ct) =>
    {
        dynamic sheet = ctx.Book.Worksheets.Item(sheetName);
        dynamic pivotTables = sheet.PivotTables;
        
        // Verify PivotTable exists in collection
        bool found = false;
        for (int i = 1; i <= pivotTables.Count; i++)
        {
            if (pivotTables.Item(i).Name == pivotTableName)
            {
                found = true;
                break;
            }
        }
        Assert.True(found, $"PivotTable '{pivotTableName}' not found in sheet '{sheetName}'");
        
        // Verify PivotTable properties
        dynamic pivotTable = pivotTables.Item(pivotTableName);
        Assert.Equal(pivotTableName, pivotTable.Name);
        Assert.True(pivotTable.PivotFields.Count >= 4); // Should have our 4 fields
        
        return ValueTask.FromResult(0);
    });
}
```

### Field Management Tests - Verify Excel COM State

```csharp
[Fact]
public async Task AddRowField_ValidField_PlacesFieldInRowAreaAndRefreshesCorrectly()
{
    // Arrange
    string testFile = await CreatePivotTableWithTestData();
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);

    // Act
    var result = await _commands.AddRowFieldAsync(batch, "TestPivot", "Region");

    // Assert - Verify Success AND Excel State
    Assert.True(result.Success);
    Assert.Equal("Region", result.FieldName);
    Assert.Equal(PivotFieldArea.Row, result.Area);
    Assert.Equal(1, result.Position); // Should be first row field
    
    // Critical: Verify actual Excel COM state
    await VerifyFieldInRowArea(batch, "TestPivot", "Region", 1);
    
    // Verify PivotTable shows row data correctly
    var pivotData = await GetPivotTableData(batch, "TestPivot");
    Assert.True(pivotData.RowHeaders.Count > 0);
    Assert.Contains("North", pivotData.RowHeaders); // Should show our test regions
    Assert.Contains("South", pivotData.RowHeaders);
}

private async Task VerifyFieldInRowArea(IExcelBatch batch, string pivotTableName, string fieldName, int expectedPosition)
{
    await batch.ExecuteAsync(async (ctx, ct) =>
    {
        dynamic pivotTable = FindPivotTable(ctx.Book, pivotTableName);
        dynamic field = pivotTable.PivotFields.Item(fieldName);
        
        // Verify orientation (1 = xlRowField)
        Assert.Equal(1, (int)field.Orientation);
        
        // Verify position in row area
        Assert.Equal(expectedPosition, (int)field.Position);
        
        // Verify field appears in RowFields collection
        dynamic rowFields = pivotTable.RowFields;
        bool foundInRowFields = false;
        for (int i = 1; i <= rowFields.Count; i++)
        {
            if (rowFields.Item(i).SourceName == fieldName)
            {
                foundInRowFields = true;
                break;
            }
        }
        Assert.True(foundInRowFields, $"Field '{fieldName}' not found in RowFields collection");
        
        return ValueTask.FromResult(0);
    });
}

[Fact]
public async Task AddValueField_NumericField_AggregatesDataCorrectly()
{
    // Arrange
    string testFile = await CreatePivotTableWithTestData();
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);
    
    // Add some structure first
    await _commands.AddRowFieldAsync(batch, "TestPivot", "Region");

    // Act
    var result = await _commands.AddValueFieldAsync(batch, "TestPivot", "Sales", AggregationFunction.Sum, "Total Sales");

    // Assert - Verify Success AND Calculation
    Assert.True(result.Success);
    Assert.Equal("Sales", result.FieldName);
    Assert.Equal("Total Sales", result.CustomName);
    Assert.Equal(PivotFieldArea.Value, result.Area);
    Assert.Equal(AggregationFunction.Sum, result.Function);
    Assert.NotNull(result.SampleValue); // Should have calculated a sample value
    
    // Critical: Verify actual aggregation in Excel
    await VerifyValueFieldAggregation(batch, "TestPivot", "Sales", AggregationFunction.Sum);
    
    // Verify PivotTable shows correct totals
    var pivotData = await GetPivotTableData(batch, "TestPivot");
    Assert.True(pivotData.Values.Count > 0);
    
    // Verify grand total calculation (North: 100+150=250, South: 200, Total: 450)
    var grandTotal = pivotData.GrandTotals["Total Sales"];
    Assert.Equal(450.0, Convert.ToDouble(grandTotal), 1); // Allow 1 decimal tolerance
}

private async Task VerifyValueFieldAggregation(IExcelBatch batch, string pivotTableName, string fieldName, AggregationFunction expectedFunction)
{
    await batch.ExecuteAsync(async (ctx, ct) =>
    {
        dynamic pivotTable = FindPivotTable(ctx.Book, pivotTableName);
        dynamic dataFields = pivotTable.DataFields;
        
        // Find our value field
        dynamic valueField = null;
        for (int i = 1; i <= dataFields.Count; i++)
        {
            var df = dataFields.Item(i);
            if (df.SourceName == fieldName)
            {
                valueField = df;
                break;
            }
        }
        
        Assert.NotNull(valueField);
        
        // Verify aggregation function (Excel COM constants)
        int expectedComFunction = expectedFunction switch
        {
            AggregationFunction.Sum => -4157,      // xlSum
            AggregationFunction.Count => -4112,    // xlCount
            AggregationFunction.Average => -4106,  // xlAverage
            _ => throw new ArgumentException($"Test not configured for {expectedFunction}")
        };
        
        Assert.Equal(expectedComFunction, (int)valueField.Function);
        
        return ValueTask.FromResult(0);
    });
}
```

### Filtering Tests - Verify Filter Effects

```csharp
[Fact]
public async Task SetFieldFilter_ValidValues_FiltersDataAndReturnsRowCounts()
{
    // Arrange
    string testFile = await CreatePivotTableWithRegionData(); // North, South, East, West
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);
    
    await _commands.AddRowFieldAsync(batch, "TestPivot", "Region");
    await _commands.AddValueFieldAsync(batch, "TestPivot", "Sales", AggregationFunction.Sum);

    // Act - Filter to show only North and South
    var result = await _commands.SetFieldFilterAsync(batch, "TestPivot", "Region", new List<string> { "North", "South" });

    // Assert - Verify Filter State AND Data Effect
    Assert.True(result.Success);
    Assert.Equal(2, result.SelectedItems.Count);
    Assert.Contains("North", result.SelectedItems);
    Assert.Contains("South", result.SelectedItems);
    Assert.Equal(4, result.AvailableItems.Count); // All 4 regions should be available
    Assert.True(result.VisibleRowCount < result.TotalRowCount); // Some rows should be hidden
    
    // Critical: Verify actual Excel filter state
    await VerifyFieldFilterState(batch, "TestPivot", "Region", new[] { "North", "South" });
    
    // Verify PivotTable only shows filtered data
    var pivotData = await GetPivotTableData(batch, "TestPivot");
    Assert.Equal(2, pivotData.RowHeaders.Count); // Only North and South
    Assert.Contains("North", pivotData.RowHeaders);
    Assert.Contains("South", pivotData.RowHeaders);
    Assert.DoesNotContain("East", pivotData.RowHeaders);
    Assert.DoesNotContain("West", pivotData.RowHeaders);
}

private async Task VerifyFieldFilterState(IExcelBatch batch, string pivotTableName, string fieldName, string[] expectedVisibleItems)
{
    await batch.ExecuteAsync(async (ctx, ct) =>
    {
        dynamic pivotTable = FindPivotTable(ctx.Book, pivotTableName);
        dynamic field = pivotTable.PivotFields.Item(fieldName);
        dynamic pivotItems = field.PivotItems;
        
        var visibleItems = new List<string>();
        var hiddenItems = new List<string>();
        
        for (int i = 1; i <= pivotItems.Count; i++)
        {
            var item = pivotItems.Item(i);
            string itemName = item.Name;
            bool isVisible = item.Visible;
            
            if (isVisible)
                visibleItems.Add(itemName);
            else
                hiddenItems.Add(itemName);
        }
        
        // Verify exactly the expected items are visible
        Assert.Equal(expectedVisibleItems.Length, visibleItems.Count);
        foreach (string expected in expectedVisibleItems)
        {
            Assert.Contains(expected, visibleItems);
        }
        
        // Verify other items are hidden
        foreach (string hidden in hiddenItems)
        {
            Assert.DoesNotContain(hidden, expectedVisibleItems);
        }
        
        return ValueTask.FromResult(0);
    });
}
```

### Data Validation Tests - Verify Calculations

```csharp
[Fact]
public async Task GetData_PivotTableWithCalculations_ReturnsCorrectAggregatedValues()
{
    // Arrange - Create PivotTable with known data for calculation verification
    string testFile = await CreatePivotTableWithKnownData();
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);
    
    // Configure PivotTable: Region in Rows, Product in Columns, Sales in Values
    await _commands.AddRowFieldAsync(batch, "TestPivot", "Region");
    await _commands.AddColumnFieldAsync(batch, "TestPivot", "Product");
    await _commands.AddValueFieldAsync(batch, "TestPivot", "Sales", AggregationFunction.Sum);

    // Act
    var result = await _commands.GetDataAsync(batch, "TestPivot");

    // Assert - Verify Structure AND Calculations
    Assert.True(result.Success);
    Assert.Equal("TestPivot", result.PivotTableName);
    Assert.True(result.Values.Count > 0);
    Assert.True(result.ColumnHeaders.Count > 0);
    Assert.True(result.RowHeaders.Count > 0);
    
    // Verify specific calculations match expected values
    // Test data: North-A: 100+150=250, North-B: 75, South-A: 200, South-B: 125
    VerifyPivotCalculation(result, "North", "A", 250.0);
    VerifyPivotCalculation(result, "North", "B", 75.0);
    VerifyPivotCalculation(result, "South", "A", 200.0);
    VerifyPivotCalculation(result, "South", "B", 125.0);
    
    // Verify grand totals
    Assert.NotNull(result.GrandTotals);
    Assert.Equal(750.0, Convert.ToDouble(result.GrandTotals["Grand Total"]), 1);
}

private void VerifyPivotCalculation(PivotTableDataResult result, string rowHeader, string columnHeader, double expectedValue)
{
    int rowIndex = result.RowHeaders.IndexOf(rowHeader);
    int colIndex = result.ColumnHeaders.IndexOf(columnHeader);
    
    Assert.True(rowIndex >= 0, $"Row header '{rowHeader}' not found");
    Assert.True(colIndex >= 0, $"Column header '{columnHeader}' not found");
    
    var actualValue = Convert.ToDouble(result.Values[rowIndex][colIndex]);
    Assert.Equal(expectedValue, actualValue, 1); // 1 decimal tolerance
}

// Test data setup helpers
private async Task<string> CreatePivotTableWithKnownData()
{
    string testFile = CreateTestWorkbook();
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);
    
    // Create precisely known test data for calculation verification
    var testData = new List<List<object?>>
    {
        new() { "Region", "Product", "Sales", "Date" },
        new() { "North", "A", 100, "2024-01-01" },
        new() { "North", "A", 150, "2024-01-02" },  // North-A total: 250
        new() { "North", "B", 75, "2024-01-03" },   // North-B total: 75
        new() { "South", "A", 200, "2024-01-04" },  // South-A total: 200
        new() { "South", "B", 125, "2024-01-05" }   // South-B total: 125
    };                                               // Grand total: 750
    
    await SetupTestData(batch, "Data", testData.ToArray());
    await batch.SaveAsync();
    
    // Create PivotTable
    var createResult = await _commands.CreateFromRangeAsync(batch, "Data", "A1:D6", "Summary", "A1", "TestPivot");
    Assert.True(createResult.Success);
    
    return testFile;
}
```

### Error Handling Tests - Verify Meaningful Error Messages

```csharp
[Fact]
public async Task AddValueField_TextFieldWithSumFunction_ReturnsInformativeError()
{
    // Arrange
    string testFile = await CreatePivotTableWithTestData();
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);

    // Act & Assert - Should fail with informative error
    var ex = await Assert.ThrowsAsync<McpException>(async () =>
        await _commands.AddValueFieldAsync(batch, "TestPivot", "Region", AggregationFunction.Sum));
    
    // Verify error message provides actionable guidance
    Assert.Contains("not valid for Text field 'Region'", ex.Message);
    Assert.Contains("Valid functions: Count", ex.Message); // Should suggest valid alternatives
}

[Fact]
public async Task AddRowField_NonexistentField_ReturnsAvailableFields()
{
    // Arrange
    string testFile = await CreatePivotTableWithTestData();
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);

    // Act & Assert
    var ex = await Assert.ThrowsAsync<McpException>(async () =>
        await _commands.AddRowFieldAsync(batch, "TestPivot", "NonexistentField"));
    
    // Verify error message lists available fields for easy correction
    Assert.Contains("Field 'NonexistentField' not found", ex.Message);
    Assert.Contains("Available fields: Region, Product, Sales, Date", ex.Message);
}
```

### Performance Tests - Verify Large Data Handling

```csharp
[Fact]
[Trait("Category", "Performance")]
public async Task CreatePivotTable_LargeDataset_CompletesWithinTimeout()
{
    // Arrange - Create large dataset (10,000 rows)
    string testFile = await CreateLargeTestDataset(10000);
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);

    var stopwatch = Stopwatch.StartNew();

    // Act
    var result = await _commands.CreateFromRangeAsync(batch, "Data", "A1:D10001", "Summary", "A1", "LargePivot");

    // Assert - Verify performance AND correctness
    stopwatch.Stop();
    Assert.True(result.Success);
    Assert.True(stopwatch.ElapsedMilliseconds < 30000); // Should complete within 30 seconds
    Assert.Equal(10000, result.SourceRowCount);
    
    // Verify PivotTable can handle the large dataset
    var infoResult = await _commands.GetInfoAsync(batch, "LargePivot");
    Assert.True(infoResult.Success);
    Assert.Equal(4, infoResult.Fields.Count);
}
```

### Basic PivotTable Creation and Configuration

```csharp
// Create PivotTable from range
await pivotCommands.CreateFromRangeAsync(batch, "Data", "A1:F1000", "Summary", "A1", "SalesPivot");

// Add fields to different areas
await pivotCommands.AddRowFieldAsync(batch, "SalesPivot", "Region");
await pivotCommands.AddRowFieldAsync(batch, "SalesPivot", "Product");
await pivotCommands.AddColumnFieldAsync(batch, "SalesPivot", "Year");
await pivotCommands.AddValueFieldAsync(batch, "SalesPivot", "Sales", AggregationFunction.Sum, "Total Sales");
await pivotCommands.AddFilterFieldAsync(batch, "SalesPivot", "Category");

// Configure formatting
await pivotCommands.SetFieldFormatAsync(batch, "SalesPivot", "Total Sales", "$#,##0");
await pivotCommands.SetLayoutAsync(batch, "SalesPivot", PivotTableLayout.Tabular);
await pivotCommands.SetStyleAsync(batch, "SalesPivot", "PivotStyleMedium2");
```

### Advanced Analysis

```csharp
// Group date field by quarters
await pivotCommands.GroupDateFieldAsync(batch, "SalesPivot", "Date", DateGrouping.Quarters);

// Create calculated field
await pivotCommands.CreateCalculatedFieldAsync(batch, "SalesPivot", "Profit Margin", 
    "= Profit / Sales * 100");

// Sort by sales values (descending)
await pivotCommands.SortFieldByValueAsync(batch, "SalesPivot", "Product", "Total Sales", 
    SortDirection.Descending);

// Filter to show only top regions
await pivotCommands.SetFieldFilterAsync(batch, "SalesPivot", "Region", 
    new List<string> { "North", "South", "East" });
```

---

## CLI Commands

### Phase 1 Commands

```bash
# === LIFECYCLE OPERATIONS ===
excelcli pivot-list <file.xlsx>
excelcli pivot-info <file.xlsx> <pivot-name>
excelcli pivot-create-from-range <file.xlsx> <src-sheet> <src-range> <dest-sheet> <dest-cell> <pivot-name>
excelcli pivot-create-from-table <file.xlsx> <table-name> <dest-sheet> <dest-cell> <pivot-name>
excelcli pivot-create-from-datamodel <file.xlsx> <connection-name> <dest-sheet> <dest-cell> <pivot-name>
excelcli pivot-delete <file.xlsx> <pivot-name>
excelcli pivot-refresh <file.xlsx> <pivot-name>
excelcli pivot-refresh-all <file.xlsx>

# === FIELD MANAGEMENT ===
excelcli pivot-list-fields <file.xlsx> <pivot-name>
excelcli pivot-add-row-field <file.xlsx> <pivot-name> <field-name> [position]
excelcli pivot-add-column-field <file.xlsx> <pivot-name> <field-name> [position]
excelcli pivot-add-value-field <file.xlsx> <pivot-name> <field-name> [function] [custom-name]
excelcli pivot-add-filter-field <file.xlsx> <pivot-name> <field-name>
excelcli pivot-remove-field <file.xlsx> <pivot-name> <field-name>
excelcli pivot-move-field <file.xlsx> <pivot-name> <field-name> <from-area> <to-area> [position]

# === FIELD CONFIGURATION ===
excelcli pivot-set-field-function <file.xlsx> <pivot-name> <field-name> <function>
excelcli pivot-set-field-name <file.xlsx> <pivot-name> <field-name> <custom-name>
excelcli pivot-set-field-format <file.xlsx> <pivot-name> <field-name> <number-format>

# === FILTERING AND SORTING ===
excelcli pivot-set-field-filter <file.xlsx> <pivot-name> <field-name> <value1,value2,...>
excelcli pivot-clear-field-filter <file.xlsx> <pivot-name> <field-name>
excelcli pivot-get-field-filter <file.xlsx> <pivot-name> <field-name>
excelcli pivot-sort-field <file.xlsx> <pivot-name> <field-name> [asc|desc]
excelcli pivot-sort-field-by-value <file.xlsx> <pivot-name> <sort-field> <value-field> [asc|desc]

# === LAYOUT AND FORMATTING ===
excelcli pivot-set-layout <file.xlsx> <pivot-name> <compact|outline|tabular>
excelcli pivot-set-style <file.xlsx> <pivot-name> <style-name>
excelcli pivot-set-grand-totals <file.xlsx> <pivot-name> <show-row-totals:true|false> <show-column-totals:true|false>
excelcli pivot-set-subtotals <file.xlsx> <pivot-name> <field-name> <show:true|false>
```

### Phase 2 Commands

```bash
# === GROUPING OPERATIONS ===
excelcli pivot-group-date-field <file.xlsx> <pivot-name> <field-name> <years|quarters|months|days> [start-date] [end-date]
excelcli pivot-group-numeric-field <file.xlsx> <pivot-name> <field-name> <start-value> <end-value> <interval>
excelcli pivot-create-custom-group <file.xlsx> <pivot-name> <field-name> <group-name> <item1,item2,...>
excelcli pivot-ungroup-field <file.xlsx> <pivot-name> <field-name>

# === CALCULATED FIELDS ===
excelcli pivot-create-calculated-field <file.xlsx> <pivot-name> <field-name> <formula>
excelcli pivot-update-calculated-field <file.xlsx> <pivot-name> <field-name> <formula>
excelcli pivot-delete-calculated-field <file.xlsx> <pivot-name> <field-name>
excelcli pivot-list-calculated-fields <file.xlsx> <pivot-name>

# === DRILL DOWN ===
excelcli pivot-drill-down <file.xlsx> <pivot-name> <target-sheet> <cell-address>

# === SLICER INTEGRATION ===
excelcli pivot-create-slicer <file.xlsx> <pivot-name> <field-name> <slicer-name> <dest-sheet> <position>
excelcli pivot-list-slicers <file.xlsx> <pivot-name>
excelcli pivot-set-slicer-selection <file.xlsx> <slicer-name> <item1,item2,...>

# === DATA SOURCE MANAGEMENT ===
excelcli pivot-change-data-source <file.xlsx> <pivot-name> <new-sheet> <new-range>
excelcli pivot-get-cache-info <file.xlsx> <pivot-name>
```

---

## MCP Tool: excel_pivottable

### Phase 1 Actions

```typescript
{
  "name": "excel_pivottable",
  "description": "Comprehensive Excel PivotTable operations - creation, field management, analysis, and formatting",
  "parameters": {
    "action": "string",
    "excelPath": "string",
    "pivotTableName": "string",
    "sourceSheet": "string",
    "sourceRange": "string",
    "tableName": "string",
    "connectionName": "string",
    "destinationSheet": "string",
    "destinationCell": "string",
    "fieldName": "string",
    "customName": "string",
    "function": "string",
    "numberFormat": "string",
    "selectedValues": "array<string>",
    "direction": "string",
    "layout": "string",
    "styleName": "string",
    "showRowTotals": "boolean",
    "showColumnTotals": "boolean",
    "showSubtotals": "boolean"
  },
  "actions": [
    // Lifecycle operations
    "list",                    // List all PivotTables
    "get-info",               // Get PivotTable details
    "create-from-range",      // Create from Excel range
    "create-from-table",      // Create from Excel Table
    "create-from-datamodel",  // Create from Data Model
    "delete",                 // Delete PivotTable
    "refresh",                // Refresh data
    "refresh-all",            // Refresh all PivotTables
    
    // Field management
    "list-fields",            // List available fields
    "add-row-field",          // Add field to Rows area
    "add-column-field",       // Add field to Columns area
    "add-value-field",        // Add field to Values area
    "add-filter-field",       // Add field to Filters area
    "remove-field",           // Remove field from any area
    "move-field",             // Move field between areas
    
    // Field configuration
    "set-field-function",     // Set aggregation function
    "set-field-name",         // Set custom field name
    "set-field-format",       // Set number format
    
    // Filtering and sorting
    "set-field-filter",       // Filter field values
    "clear-field-filter",     // Clear field filter
    "get-field-filter",       // Get filter state
    "sort-field",             // Sort field
    "sort-field-by-value",    // Sort by value field
    
    // Layout and formatting
    "set-layout",             // Set PivotTable layout
    "set-style",              // Apply PivotTable style
    "set-grand-totals",       // Configure grand totals
    "set-subtotals"           // Configure subtotals
  ]
}
```

---

## Relationship to Existing Commands

### Clear Separation of Concerns

**PivotTableCommands** (New):
- Create, configure, and manage PivotTables
- Field layout and aggregation
- PivotTable-specific filtering and sorting
- PivotCache management

**TableCommands** (Existing):
- Excel Table (ListObject) operations
- Structured data with headers
- Table filtering and styling
- Can be **data source** for PivotTables

**RangeCommands** (Existing):
- Direct cell/range data operations
- Can be **data source** for PivotTables
- Can read **output** from PivotTables

**DataModelCommands** (Existing):
- Power Pivot Data Model
- DAX measures and relationships
- Can be **data source** for PivotTables

### Workflow Integration

```csharp
// 1. Create and populate data source (TableCommands)
await tableCommands.CreateAsync(batch, "Data", "SalesTable", "A1:F1000", true);

// 2. Create PivotTable from table (PivotTableCommands)
await pivotCommands.CreateFromTableAsync(batch, "SalesTable", "Summary", "A1", "SalesPivot");

// 3. Configure PivotTable fields (PivotTableCommands)
await pivotCommands.AddRowFieldAsync(batch, "SalesPivot", "Region");
await pivotCommands.AddValueFieldAsync(batch, "SalesPivot", "Sales", AggregationFunction.Sum);

// 4. Read PivotTable results (RangeCommands)
var results = await rangeCommands.GetValuesAsync(batch, "Summary", "A1:D20");
```

---

## Performance Considerations

### PivotCache Management

1. **Reuse Caches**: Multiple PivotTables can share same PivotCache
2. **Refresh Strategy**: Batch refresh operations when possible
3. **Memory Usage**: Large datasets may require cache optimization

### Field Operations

1. **Batch Configuration**: Configure multiple fields before refresh
2. **Validation**: Check field existence before operations
3. **Error Handling**: Graceful handling of invalid field configurations

---

## Security Considerations

### Data Source Access

1. **Permission Validation**: Ensure user has access to data sources
2. **Connection Security**: Validate external data connections
3. **Range Validation**: Verify source ranges exist and are accessible

---

## Success Criteria

### Phase 1 (MVP)
- [ ] Complete PivotTable lifecycle management
- [ ] Field management for all areas (Rows, Columns, Values, Filters)
- [ ] Basic filtering and sorting capabilities
- [ ] Layout and formatting options
- [ ] Integration with existing TableCommands and RangeCommands
- [ ] Comprehensive CLI commands
- [ ] MCP Server tool with all actions
- [ ] Complete test coverage

### Phase 2 (Advanced)
- [ ] Date and numeric grouping
- [ ] Calculated fields
- [ ] Drill-down functionality
- [ ] Slicer integration
- [ ] Advanced data source management

---

## Implementation Timeline

**Phase 1 (Core Operations)**: 1-2 weeks
- PivotTable lifecycle and basic field management
- CLI commands and MCP tool
- Integration tests

**Phase 2 (Advanced Features)**: 1 week
- Grouping, calculated fields, slicers
- Advanced CLI commands
- Extended test coverage

**Total Implementation**: 2-3 weeks for complete PivotTable functionality

---

## Open Questions

1. **Slicer Management**: Should slicers be part of PivotTableCommands or separate SlicerCommands?
2. **Multiple PivotTables**: How to handle operations affecting multiple PivotTables sharing same cache?
3. **Data Model Integration**: Should advanced Data Model PivotTables use TOM API like DataModelCommands?
4. **Chart Integration**: Should PivotChart creation be included or handled separately?

**Recommended Answers**:
1. **Include in PivotTableCommands** - Slicers are primarily PivotTable features
2. **Individual targeting** - Operations target specific PivotTable, let Excel handle cache sharing
3. **Use Excel COM** - Keep consistent with PivotTable object model, use TOM only for Data Model-specific operations
4. **Separate concern** - PivotCharts could be future ChartCommands (different abstraction level)

---

## MCP Server Implementation (LLM-Optimized Design)

### ExcelPivotTableTool Design Philosophy

**LLM-Friendly Actions**: Each action should be intuitive and provide rich context for AI reasoning.

```csharp
[McpServerTool]
public async Task<string> ExcelPivotTable(
    string action,
    string excelPath,
    string? pivotTableName = null,
    string? sourceSheet = null,
    string? sourceRange = null,
    string? targetSheet = null,
    string? targetCell = null,
    string? fieldName = null,
    string? customName = null,
    string? aggregationFunction = null,
    object? filterValues = null,
    object? sortColumns = null,
    string? layoutTemplate = null)
{
    return action.ToLowerInvariant() switch
    {
        // Core lifecycle (LLM can create, explore, remove)
        "create-from-range" => await CreateFromRange(excelPath, sourceSheet!, sourceRange!, targetSheet!, targetCell!, pivotTableName!),
        "create-from-table" => await CreateFromTable(excelPath, sourceSheet!, tableOrRangeName!, targetSheet!, targetCell!, pivotTableName!),
        "delete" => await DeletePivotTable(excelPath, pivotTableName!),
        "list" => ListPivotTables(excelPath),
        
        // Discovery (LLM can understand structure and guide configuration)
        "get-info" => await GetInfo(excelPath, pivotTableName!),
        "get-fields" => await GetFields(excelPath, pivotTableName!),
        "get-data" => await GetData(excelPath, pivotTableName!),
        "get-layout" => await GetLayout(excelPath, pivotTableName!),
        
        // Field management (LLM can build analysis step-by-step)
        "add-row-field" => await AddRowField(excelPath, pivotTableName!, fieldName!),
        "add-column-field" => await AddColumnField(excelPath, pivotTableName!, fieldName!),
        "add-value-field" => await AddValueField(excelPath, pivotTableName!, fieldName!, aggregationFunction, customName),
        "add-page-field" => await AddPageField(excelPath, pivotTableName!, fieldName!),
        "remove-field" => await RemoveField(excelPath, pivotTableName!, fieldName!),
        "move-field" => await MoveField(excelPath, pivotTableName!, fieldName!, newArea!, newPosition),
        
        // Data manipulation (LLM can filter and sort interactively)
        "set-field-filter" => await SetFieldFilter(excelPath, pivotTableName!, fieldName!, filterValues!),
        "clear-field-filter" => await ClearFieldFilter(excelPath, pivotTableName!, fieldName!),
        "clear-all-filters" => await ClearAllFilters(excelPath, pivotTableName!),
        "sort-field" => await SortField(excelPath, pivotTableName!, fieldName!, sortOrder!),
        
        // Layout and formatting (LLM can apply templates and styles)
        "apply-layout-template" => await ApplyLayoutTemplate(excelPath, pivotTableName!, layoutTemplate!),
        "refresh" => await Refresh(excelPath, pivotTableName!),
        
        _ => ThrowUnknownAction(action, validActions)
    };
}
```

### LLM Workflow Examples

**Scenario 1: LLM creates analysis from scratch**
```typescript
// Step 1: LLM explores available data
excel_pivottable({ 
    action: "create-from-range", 
    excelPath: "sales.xlsx", 
    sourceSheet: "RawData", 
    sourceRange: "A1:F1000",
    targetSheet: "Analysis",
    targetCell: "A1",
    pivotTableName: "SalesAnalysis"
})
// Returns: { success: true, availableFields: ["Region", "Product", "Sales", "Date", "Salesperson", "Category"], numericFields: ["Sales"], dateFields: ["Date"] }

// Step 2: LLM builds row structure
excel_pivottable({ 
    action: "add-row-field", 
    excelPath: "sales.xlsx", 
    pivotTableName: "SalesAnalysis", 
    fieldName: "Region"
})
// Returns: { success: true, fieldName: "Region", area: "Row", position: 1, uniqueValues: ["North", "South", "East", "West"] }

// Step 3: LLM adds analysis dimension
excel_pivottable({ 
    action: "add-column-field", 
    excelPath: "sales.xlsx", 
    pivotTableName: "SalesAnalysis", 
    fieldName: "Product"
})
// Returns: { success: true, fieldName: "Product", area: "Column", position: 1, uniqueValues: ["Product A", "Product B", "Product C"] }

// Step 4: LLM adds metrics
excel_pivottable({ 
    action: "add-value-field", 
    excelPath: "sales.xlsx", 
    pivotTableName: "SalesAnalysis", 
    fieldName: "Sales",
    aggregationFunction: "Sum",
    customName: "Total Sales"
})
// Returns: { success: true, fieldName: "Sales", customName: "Total Sales", function: "Sum", sampleValue: 125000.0 }

// Step 5: LLM applies filtering for focused analysis
excel_pivottable({ 
    action: "set-field-filter", 
    excelPath: "sales.xlsx", 
    pivotTableName: "SalesAnalysis", 
    fieldName: "Region",
    filterValues: ["North", "South"]
})
// Returns: { success: true, selectedItems: ["North", "South"], visibleRowCount: 250, totalRowCount: 500 }
```

### Rich Result Types for LLM Consumption

Each action returns detailed information that helps LLMs make informed decisions:

```csharp
// Create operations return field analysis
public class PivotTableCreateResult : OperationResult
{
    public string PivotTableName { get; set; } = string.Empty;
    public string SourceRange { get; set; } = string.Empty;
    public string TargetLocation { get; set; } = string.Empty;
    public int SourceRowCount { get; set; }
    public List<string> AvailableFields { get; set; } = new();
    public List<string> NumericFields { get; set; } = new();    // LLM can suggest Sum, Average
    public List<string> DateFields { get; set; } = new();      // LLM can suggest grouping
    public List<string> TextFields { get; set; } = new();      // LLM can suggest Count
    public Dictionary<string, int> FieldValueCounts { get; set; } = new(); // LLM can assess cardinality
}

// Field operations return impact analysis
public class PivotFieldResult : OperationResult
{
    public string FieldName { get; set; } = string.Empty;
    public PivotFieldArea Area { get; set; }
    public int Position { get; set; }
    public List<string> UniqueValues { get; set; } = new();    // LLM can understand filter options
    public int ValueCount { get; set; }                        // LLM can assess performance impact
    public object? SampleValue { get; set; }                   // LLM can verify data types
    public List<string> SuggestedNextActions { get; set; } = new(); // Guide LLM workflow
}

// Filter operations return visibility impact
public class PivotFilterResult : OperationResult
{
    public string FieldName { get; set; } = string.Empty;
    public List<string> SelectedItems { get; set; } = new();
    public List<string> AvailableItems { get; set; } = new();
    public int VisibleRowCount { get; set; }                   // LLM can understand filter impact
    public int TotalRowCount { get; set; }
    public double FilteredPercentage => TotalRowCount > 0 ? (double)VisibleRowCount / TotalRowCount * 100 : 0;
}

// Data operations return structured analysis results
public class PivotTableDataResult : OperationResult
{
    public string PivotTableName { get; set; } = string.Empty;
    public List<string> RowHeaders { get; set; } = new();
    public List<string> ColumnHeaders { get; set; } = new();
    public List<List<object?>> Values { get; set; } = new();   // LLM can analyze patterns
    public Dictionary<string, object?> GrandTotals { get; set; } = new();
    public Dictionary<string, object?> RowTotals { get; set; } = new();
    public Dictionary<string, object?> ColumnTotals { get; set; } = new();
    public DateTime LastRefresh { get; set; }
    public string DataSummary { get; set; } = string.Empty;    // Human-readable summary for LLM
}
```

### Error Handling for LLMs

Provide actionable error messages that help LLMs correct issues:

```csharp
private async Task<string> AddValueField(string excelPath, string pivotTableName, string fieldName, string? aggregationFunction, string? customName)
{
    try
    {
        var result = await _commands.AddValueFieldAsync(batch, pivotTableName, fieldName, function, customName);
        return JsonSerializer.Serialize(result, JsonOptions);
    }
    catch (InvalidFieldTypeException ex)
    {
        // LLM-friendly error with suggestions
        var error = new
        {
            success = false,
            error = "invalid_field_type",
            message = ex.Message,
            fieldName = fieldName,
            detectedType = ex.FieldType,
            validFunctions = ex.ValidFunctions,  // ["Count"] for text fields
            suggestion = $"For {ex.FieldType} fields, try using 'Count' instead of '{aggregationFunction}'"
        };
        return JsonSerializer.Serialize(error, JsonOptions);
    }
    catch (FieldNotFoundException ex)
    {
        var error = new
        {
            success = false,
            error = "field_not_found",
            message = ex.Message,
            requestedField = fieldName,
            availableFields = ex.AvailableFields,
            suggestion = ex.AvailableFields.Count > 0 ? $"Did you mean: {ex.AvailableFields.First()}?" : "Check field names with 'get-fields' action"
        };
        return JsonSerializer.Serialize(error, JsonOptions);
    }
}
```

### Layout Templates for LLM Quick Start

```csharp
public static class PivotLayoutTemplates
{
    public static PivotLayoutTemplate SalesAnalysis => new()
    {
        Name = "Sales Analysis",
        Description = "Region/Product cross-analysis with sales metrics",
        RowFields = new[] { "Region" },
        ColumnFields = new[] { "Product" },
        ValueFields = new[] 
        { 
            new ValueFieldTemplate("Sales", AggregationFunction.Sum, "Total Sales"),
            new ValueFieldTemplate("Sales", AggregationFunction.Count, "Transaction Count")
        },
        PageFields = new[] { "Date" }, // For date filtering
        DefaultFilters = new Dictionary<string, List<string>>(),
        Style = "TableStyleMedium9"
    };
    
    public static PivotLayoutTemplate TimeSeriesAnalysis => new()
    {
        Name = "Time Series Analysis",
        Description = "Date-based trending with metrics over time",
        RowFields = new[] { "Date" },
        ColumnFields = new[] { "Category" },
        ValueFields = new[] 
        { 
            new ValueFieldTemplate("Amount", AggregationFunction.Sum, "Total Amount"),
            new ValueFieldTemplate("Amount", AggregationFunction.Average, "Average Amount")
        },
        GroupDateFields = true, // Group dates by month/quarter
        Style = "TableStyleLight16"
    };
}
```
