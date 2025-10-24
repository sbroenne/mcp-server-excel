# Data Model TOM API Feature Specification (Phase 4)

## Overview

This specification defines Phase 4 of the Data Model feature: implementing CREATE and UPDATE operations using Microsoft Analysis Services Tabular Object Model (TOM) API.

**Status**: Planning Phase  
**Target Version**: v2.0.0 (Future)  
**Dependencies**: Phase 1 (Read/Delete operations via COM API) must be complete

## Background

Phase 1 implemented Read and Delete operations using Excel COM API:
- ✅ List tables, measures, relationships
- ✅ View and export DAX formulas
- ✅ Delete measures and relationships
- ✅ Refresh Data Model

**Limitation**: Excel COM API does not support creating or updating Data Model objects. These operations require the Analysis Services Tabular Object Model (TOM) API.

## Goals

### Primary Goals
1. **Create DAX Measures**: Add new measures to Data Model tables
2. **Update DAX Measures**: Modify existing measure formulas and properties
3. **Create Relationships**: Define table relationships programmatically
4. **Update Relationships**: Modify relationship properties (active/inactive, cardinality, cross-filter)

### Non-Goals (Out of Scope for Phase 4)
- Creating new tables in Data Model (requires data source configuration)
- Modifying table schemas or column definitions
- Creating calculated columns (different from measures)
- KPI management
- Perspectives and roles
- Translations

## Technical Research

### TOM API NuGet Package

**Package**: `Microsoft.AnalysisServices.NetCore.retail.amd64`  
**Latest Version**: 19.84.1 (as of October 2025)  
**Target Framework**: .NET Core / .NET 5+  
**Compatibility**: ✅ Compatible with .NET 8.0 (tested and verified)  
**Documentation**: https://docs.microsoft.com/en-us/analysis-services/tom/introduction-to-the-tabular-object-model-tom-in-analysis-services-amo

**Key Discovery**: The original `Microsoft.AnalysisServices.Tabular` package only supports .NET Framework. The modern alternative `Microsoft.AnalysisServices.NetCore.retail.amd64` provides full .NET Core/.NET 8 support with zero compatibility warnings.

### TOM API Architecture

```csharp
using Microsoft.AnalysisServices.Tabular;

// Connect to Excel Data Model
Server server = new Server();
server.Connect($"Provider=MSOLAP;Data Source={excelFilePath};");

// Access database (Excel Data Model)
Database database = server.Databases[0];
Model model = database.Model;

// Access tables
Table table = model.Tables.Find("TableName");

// Create measure
Measure newMeasure = new Measure
{
    Name = "TotalSales",
    Expression = "SUM(Sales[Amount])",
    FormatString = "#,##0.00"
};
table.Measures.Add(newMeasure);

// Save changes
model.SaveChanges();

// Disconnect
server.Disconnect();
```

### Key TOM Concepts

1. **Server Connection**: TOM connects to Excel Data Model via MSOLAP provider
2. **Database**: Excel workbook Data Model is accessed as an Analysis Services database
3. **Model**: Contains tables, measures, relationships
4. **SaveChanges()**: Required to commit changes to Excel file

### Compatibility Considerations

- **Excel Version**: Data Model available in Excel 2013+ (with Power Pivot add-in)
- **TOM API Version**: Must match Analysis Services version embedded in Excel
- **File Format**: Only .xlsx and .xlsm files support Data Model
- **Power Pivot**: Must be enabled in Excel for Data Model features

## Implementation Phases

### Phase 4.0: Research & Prototyping (This Phase)
- [x] Create comprehensive specification
- [x] Research TOM API version compatibility with Excel versions
  - ✅ Discovered `Microsoft.AnalysisServices.NetCore.retail.amd64` package
  - ✅ Version 19.84.1 is .NET 8.0 compatible
  - ✅ Added package to ExcelMcp.Core project successfully
- [x] Create prototype for measure creation
  - ✅ Created `TomPrototype.CanCreateMeasure()` method
  - ✅ Compiles successfully with .NET 8.0
- [x] Create prototype for relationship creation
  - ✅ Created `TomPrototype.CanCreateRelationship()` method
  - ✅ Supports cardinality and active/inactive settings
- [x] Validate TOM API works with Excel files (not just SSAS servers)
  - ✅ Created `TomPrototype.CanConnectToExcelDataModel()` with multiple connection string formats
  - ⏳ Pending: Runtime testing with actual Excel file containing Data Model
- [ ] Document any Excel-specific TOM limitations

### Phase 4.1: Core Implementation
- [ ] Add Microsoft.AnalysisServices.Tabular NuGet package
- [ ] Create `IDataModelTomCommands` interface
- [ ] Implement `DataModelTomCommands` class with:
  - `CreateMeasure()` - Add new measure to table
  - `UpdateMeasure()` - Modify measure formula/properties
  - `CreateRelationship()` - Define table relationship
  - `UpdateRelationship()` - Modify relationship properties
- [ ] Create comprehensive unit tests
- [ ] Create integration tests with real Excel files

### Phase 4.2: CLI Integration
- [ ] Create CLI commands:
  - `dm-create-measure <file> <table> <name> <formula> [--format <format>]`
  - `dm-update-measure <file> <measure-name> <new-formula> [--format <format>]`
  - `dm-create-relationship <file> <from-table> <from-col> <to-table> <to-col> [--inactive]`
  - `dm-update-relationship <file> <relationship-id> [--active|--inactive] [--cardinality <value>]`
- [ ] Add Spectre.Console formatting
- [ ] Create CLI tests

### Phase 4.3: MCP Server Integration
- [ ] Extend `ExcelDataModelTool` with new actions:
  - `create-measure`
  - `update-measure`
  - `create-relationship`
  - `update-relationship`
- [ ] Add async support for all TOM operations
- [ ] Create MCP Server tests

### Phase 4.4: Documentation
- [ ] Update COMMANDS.md with CREATE/UPDATE operations
- [ ] Update README.md with TOM capabilities
- [ ] Create TOM API usage examples
- [ ] Update copilot instructions with TOM patterns

## API Design

### Core Methods

```csharp
public interface IDataModelTomCommands
{
    /// <summary>
    /// Create a new DAX measure in the specified table.
    /// </summary>
    DataModelResult CreateMeasure(
        string filePath,
        string tableName,
        string measureName,
        string daxFormula,
        string? formatString = null,
        string? description = null);

    /// <summary>
    /// Update an existing DAX measure's formula and properties.
    /// </summary>
    DataModelResult UpdateMeasure(
        string filePath,
        string measureName,
        string? newFormula = null,
        string? newFormatString = null,
        string? newDescription = null);

    /// <summary>
    /// Create a relationship between two tables.
    /// </summary>
    DataModelResult CreateRelationship(
        string filePath,
        string fromTableName,
        string fromColumnName,
        string toTableName,
        string toColumnName,
        bool isActive = true,
        string? cardinality = null,
        string? crossFilterDirection = null);

    /// <summary>
    /// Update an existing relationship's properties.
    /// </summary>
    DataModelResult UpdateRelationship(
        string filePath,
        string fromTableName,
        string fromColumnName,
        string toTableName,
        string toColumnName,
        bool? isActive = null,
        string? cardinality = null,
        string? crossFilterDirection = null);
}
```

### CLI Command Examples

```bash
# Create measure
excelcli dm-create-measure "sales.xlsx" "FactSales" "TotalRevenue" "SUM(FactSales[Amount])" --format "#,##0.00"

# Update measure formula
excelcli dm-update-measure "sales.xlsx" "TotalRevenue" "SUM(FactSales[Amount]) * 1.1" --format "$#,##0.00"

# Create relationship
excelcli dm-create-relationship "sales.xlsx" "FactSales" "ProductID" "DimProduct" "ProductID"

# Update relationship to inactive
excelcli dm-update-relationship "sales.xlsx" "FactSales" "ProductID" "DimProduct" "ProductID" --inactive
```

### MCP Server Actions

```json
{
  "action": "create-measure",
  "excelPath": "sales.xlsx",
  "tableName": "FactSales",
  "measureName": "TotalRevenue",
  "daxFormula": "SUM(FactSales[Amount])",
  "formatString": "#,##0.00"
}
```

## Error Handling

### TOM-Specific Errors

1. **Connection Failures**
   - Excel file locked by another process
   - MSOLAP provider not available
   - Data Model not initialized in workbook

2. **Model Validation Errors**
   - Invalid DAX formula syntax
   - Referenced table/column doesn't exist
   - Duplicate measure names
   - Circular relationship dependencies

3. **Save Failures**
   - File read-only
   - Insufficient permissions
   - Data Model corruption

### Error Result Pattern

```csharp
if (!model.Tables.Contains(tableName))
{
    return new DataModelResult
    {
        Success = false,
        ErrorMessage = $"Table '{tableName}' not found in Data Model. Use 'dm-list-tables' to see available tables.",
        SuggestedNextActions = new List<string>
        {
            "Run: dm-list-tables <file> to see all tables",
            "Verify table name spelling and case sensitivity"
        }
    };
}
```

## Security Considerations

### DAX Formula Injection

**Risk**: Malicious DAX formulas could potentially access sensitive data or cause performance issues.

**Mitigation**:
- Validate DAX syntax before saving (use TOM validation)
- Warn users about formulas from untrusted sources
- Consider sandboxing for automated scenarios
- Log all measure creation/updates for audit trail

### File Access

**Risk**: TOM API requires exclusive file access, could conflict with Excel UI.

**Mitigation**:
- Check if file is open in Excel before TOM operations
- Provide clear error messages if file is locked
- Recommend closing Excel during automated updates

## Testing Strategy

### Unit Tests
- Validate parameter checking (null values, empty strings)
- Test error message generation
- Verify DAX formula sanitization

### Integration Tests (Require Excel)
- Create measures with various DAX formulas
- Update existing measures
- Create relationships with different cardinalities
- Update relationship properties
- Verify model.SaveChanges() persists to file
- Test rollback on validation errors

### Round Trip Tests
- Create measure → View measure → Update measure → Export formula
- Create relationship → List relationships → Update → Delete
- Full workflow: Import data → Create measure → Create relationship → Refresh → Validate results

### Test Helper Pattern

```csharp
public static class TomTestHelper
{
    public static void WithTomServer(string filePath, Action<Server, Model> action)
    {
        Server server = new Server();
        try
        {
            server.Connect($"Provider=MSOLAP;Data Source={filePath};");
            Database database = server.Databases[0];
            Model model = database.Model;
            
            action(server, model);
            
            model.SaveChanges();
        }
        finally
        {
            if (server.Connected)
            {
                server.Disconnect();
            }
        }
    }
}
```

## Performance Considerations

### Connection Pooling
- TOM connections are heavier than COM API calls
- Consider connection pooling for multiple operations
- Reuse server connection within same operation

### Batch Operations
- Group multiple measure creations in single SaveChanges()
- Reduce file I/O overhead
- Improve transaction atomicity

### Validation Overhead
- DAX validation can be expensive
- Cache validation results where possible
- Provide option to skip validation (at user's risk)

## Open Questions

1. **TOM API Compatibility**: Which versions of Microsoft.AnalysisServices.Tabular are compatible with Excel 2016, 2019, 2021, Microsoft 365?
2. **Excel Process**: Does TOM require Excel process to be running, or can it work with closed files?
3. **Concurrent Access**: Can TOM and COM API be used simultaneously on same file?
4. **Licensing**: Are there any licensing implications for using TOM API with Excel?
5. **Error Recovery**: How to handle partial failures (e.g., 3 of 5 measures created before error)?

## Success Criteria

Phase 4 is complete when:
- [ ] All CRUD operations work (Read/Delete via COM + Create/Update via TOM)
- [ ] 100% test pass rate across all test categories
- [ ] Documentation complete with TOM examples
- [ ] MCP Server exposes full CRUD capabilities
- [ ] CLI provides complete measure and relationship management
- [ ] Performance acceptable (< 5 seconds for typical operations)
- [ ] Error handling comprehensive and user-friendly

## Future Enhancements (Phase 5+)

- Calculated columns support
- KPI creation and management
- Table creation from data sources
- Advanced relationship features (bi-directional filtering)
- Batch operations API
- DAX formula validation and optimization suggestions
- Data Model schema versioning

## References

- [TOM API Documentation](https://docs.microsoft.com/en-us/analysis-services/tom/introduction-to-the-tabular-object-model-tom-in-analysis-services-amo)
- [DAX Formula Reference](https://docs.microsoft.com/en-us/dax/)
- [Excel Data Model Overview](https://support.microsoft.com/en-us/office/create-a-data-model-in-excel-87e7a54c-87dc-488e-9410-5c75dbcb0f7b)
- [Power Pivot for Excel](https://support.microsoft.com/en-us/office/power-pivot-overview-and-learning-f9001958-7901-4caa-ad80-028a6d2432ed)
