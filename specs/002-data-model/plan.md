# Implementation Plan: Data Model and DAX Management

**Feature**: Data Model and DAX Management  
**Branch**: `002-data-model`  
**Status**: ✅ **IMPLEMENTED**  
**Last Updated**: 2025-11-10

## Implementation Status

### Phase 1: Excel COM API Layer (✅ COMPLETE)
- ✅ Read operations via `Workbook.Model`
- ✅ List tables, measures, relationships
- ✅ Export DAX formulas to files
- ✅ Get model metadata

### Phase 2: TOM API Layer (✅ COMPLETE)
- ✅ Write operations via TOM API
- ✅ Create/Update/Delete DAX measures
- ✅ Create/Update/Delete relationships
- ✅ DAX validation on create/update

### Phase 3: Advanced Features (🔜 PLANNED)
- ❌ Calculated columns
- ❌ Hierarchies and perspectives
- ❌ Row-level security (RLS)
- ❌ Advanced DAX validation and testing

## Architecture Overview

### Two-Layer API Strategy

**Layer 1 - Excel COM API (Read Operations)**
```
Workbook.Model
    ├── ModelTables[1..n]
    │   ├── Name, SourceName, RecordCount
    │   └── ModelTableColumns[1..n]
    ├── ModelMeasures[1..n]  (READ-ONLY formulas)
    │   ├── Name, AssociatedTable
    │   └── Formula (via TOM API required for write)
    └── ModelRelationships[1..n]
        ├── ForeignKeyColumn, PrimaryKeyColumn
        └── Active, CrossFilterDirection
```

**Layer 2 - Tabular Object Model (Write Operations)**
```
Server (MSOLAP Provider)
    └── Database (Excel File)
        ├── Model
        │   ├── Tables[1..n]
        │   │   ├── Measures[1..n]  (CRUD via TOM)
        │   │   ├── Columns[1..n]
        │   │   └── Partitions[1..n]
        │   └── Relationships[1..n]  (CRUD via TOM)
        └── SaveChanges() → Excel file
```

### Component Structure

```
src/ExcelMcp.Core/Commands/DataModel/
├── IDataModelCommands.cs           # Interface (15 methods)
├── DataModelCommands.cs            # Partial class (constructor, DI)
├── DataModelCommands.Read.cs       # Excel COM read operations
├── DataModelCommands.Write.cs      # TOM API write operations
└── DataModelHelpers.cs             # Shared utilities
```

## Technical Architecture

### Technology Stack

- **Runtime**: .NET 9.0, Windows-only
- **COM Interop**: `dynamic` types with late binding
- **TOM API**: `Microsoft.AnalysisServices.NetCore.retail.amd64` NuGet package
- **Connection**: `Data Source={excelFile};Provider=MSOLAP`
- **Batch API**: `IExcelBatch` for exclusive workbook access
- **Serialization**: `System.Text.Json` for model metadata
- **Testing**: xUnit with Excel COM integration tests

### Key Design Decisions

#### Decision 1: Two-Layer API Approach
**Rationale**: Excel COM API is read-only for most Data Model operations. TOM API provides full CRUD but requires separate connection.

**Trade-offs**:
- ✅ Excel COM for fast read operations (no extra dependencies)
- ✅ TOM API for reliable write operations (proper DAX validation)
- ⚠️ Two connection patterns to maintain
- ⚠️ Large NuGet dependency (AMO.dll)

#### Decision 2: TOM Connection via MSOLAP Provider
**Implementation**:
```csharp
string connectionString = $"Data Source={excelFilePath};Provider=MSOLAP";
using var server = new Server();
server.Connect(connectionString);
Database database = server.Databases[0];
Model model = database.Model;

// CRUD operations...
model.SaveChanges(); // Commits to Excel file
```

**Why**: MSOLAP provider is the official connection method for embedded Excel Data Models.

#### Decision 3: Measure Export Format with Metadata
**Implementation**:
```dax
-- Measure: Total Sales
-- Table: Sales
-- Description: Sum of all sales amounts
-- Format: $#,##0.00
-- Created: 2025-01-10

Total Sales := 
SUM(Sales[Amount])
```

**Why**: Metadata headers make .dax files self-documenting and git-friendly.

#### Decision 4: Synchronous Refresh with Timeout Support
**Implementation**:
```csharp
public async Task<OperationResult> RefreshAsync(
    IExcelBatch batch, 
    CancellationToken cancellationToken = default)
{
    return await batch.ExecuteAsync((ctx, ct) => {
        ctx.Book.Model.Refresh();
        return new OperationResult { Success = true };
    }, 
    cancellationToken, 
    timeout: TimeSpan.FromMinutes(5)); // Extended timeout for refresh
}
```

**Why**: Data Model refresh can take several minutes with large datasets. Default 2-min timeout is insufficient.

#### Decision 5: No Calculated Columns (Yet)
**Rationale**: Calculated columns require TOM API and add complexity (DAX evaluation timing, memory impact).

**Current Status**: Deferred to Phase 3 - measures are higher priority for development workflows.

### Security Considerations

- **No Credential Storage**: Data Model connections use existing Excel workbook sources
- **DAX Injection Risk**: Low - TOM API validates syntax, no dynamic SQL
- **File Access**: Validated via `FileAccessValidator` before TOM connection
- **COM Object Cleanup**: Guaranteed via try/finally and `ComUtilities.Release()`

### Performance Optimizations

1. **Batch API Usage**: All operations use `IExcelBatch` to minimize workbook open/close cycles
2. **Bulk Export**: Export all measures in single TOM connection session
3. **Selective Refresh**: Refresh individual tables vs entire model (future enhancement)
4. **TOM Connection Reuse**: Single TOM connection per batch operation

## Testing Strategy

### Integration Test Coverage

**File**: `tests/ExcelMcp.Core.Tests/Commands/DataModelCommandsTests.cs`

**Test Categories**:
1. **List Operations**:
   - List tables (empty model, multiple tables, Power Query sources)
   - List measures (all, filtered by table)
   - List columns (all data types)
   - List relationships (active/inactive)

2. **Export Operations**:
   - Export single measure to .dax file
   - Export all measures to directory
   - Verify metadata headers in output

3. **Create Operations**:
   - Create measure via TOM API
   - Create relationship with cardinality
   - Validate measure appears in Excel

4. **Update Operations**:
   - Update measure formula
   - Update relationship (activate/deactivate)
   - Verify changes persist after save/reopen

5. **Delete Operations**:
   - Delete measure
   - Delete relationship
   - Verify removal from model

6. **Refresh Operations**:
   - Refresh model with default timeout
   - Refresh model with extended timeout (5 min)
   - Handle timeout errors gracefully

### Test Data Requirements

- **Empty Data Model**: Workbook with no tables/measures
- **Simple Model**: 2 tables (Sales, Products), 3 measures, 1 relationship
- **Complex Model**: 5+ tables, 10+ measures, multiple relationships
- **Power Query Source**: Tables loaded via Power Query (test refresh)

### Manual Testing

1. Open test workbook in Excel → Power Pivot → Verify tables/measures match List output
2. Create measure via CLI → Refresh Excel → Verify measure appears
3. Export all measures → Open .dax files → Verify metadata and formulas
4. Update measure formula → Reopen workbook → Verify change persisted
5. Delete measure → Refresh Power Pivot UI → Verify removed

## Known Limitations

### Current Limitations
- **Calculated Columns**: Not implemented (requires TOM schema enhancements)
- **Hierarchies**: Not implemented (future TOM feature)
- **Perspectives**: Not implemented (advanced TOM feature)
- **Row-Level Security**: Not implemented (requires DAX role management)
- **DAX IntelliSense**: Not available programmatically (Excel-only feature)

### Excel COM API Limitations
- **Read-Only Measures**: Cannot modify DAX formulas via Excel COM
- **No Validation**: Excel COM doesn't validate DAX syntax
- **Limited Metadata**: Format strings require TOM API access

### TOM API Limitations
- **File Locking**: TOM connection requires exclusive access (handled by Batch API)
- **Large Models**: Models >1GB may hit memory limits during TOM operations
- **Version Compatibility**: TOM API version must match Analysis Services embedded version

## Deployment Considerations

### NuGet Dependencies
- `Microsoft.AnalysisServices.NetCore.retail.amd64` (130+ MB)
- Platform-specific build (amd64 only, Windows-only)

### Runtime Requirements
- Excel installed (for embedded Analysis Services engine)
- .NET 9.0 runtime
- Windows OS (COM + MSOLAP provider)

### CI/CD Impact
- Integration tests require Excel COM + Analysis Services
- Tests run on Azure self-hosted runner (see `AZURE_SELFHOSTED_RUNNER_SETUP.md`)
- Cannot run in GitHub Actions hosted runners (no Excel)

## Migration Notes

### Breaking Changes
None - this is a new feature in v1.0.0

### Future Compatibility
- TOM API version may need updates for newer Excel versions
- DAX language features depend on embedded Analysis Services version
- Calculated columns/hierarchies will require schema migrations

## Related Documentation

- **Original Spec**: `specs/DATA-MODEL-DAX-FEATURE-SPEC.md`
- **TOM API Reference**: https://docs.microsoft.com/analysis-services/tom/
- **DAX Guide**: https://dax.guide
- **Testing Strategy**: `.github/instructions/testing-strategy.instructions.md`
- **Excel COM Patterns**: `.github/instructions/excel-com-interop.instructions.md`
