# Feature Specification: Data Model and DAX Management

**Feature Branch**: `002-data-model`  
**Created**: 2024-01-10  
**Status**: ✅ **IMPLEMENTED** (Excel COM API Layer Only)  
**Last Updated**: 2025-11-10

## Implementation Status

This feature is **PARTIALLY IMPLEMENTED** - Excel COM API operations only. TOM (Tabular Object Model) operations are **PLANNED** but not yet implemented.

**✅ Implemented (Excel COM API - Read Operations):**
- ✅ List tables (`ListTablesAsync`)
- ✅ List columns (`ListColumnsAsync`)
- ✅ Get table details (`GetTableAsync`)
- ✅ Get model info (`GetInfoAsync`)
- ✅ List measures (`ListMeasuresAsync`)
- ✅ Get measure details (`GetAsync`)
- ✅ Export measure DAX (`ExportMeasureAsync`)
- ✅ List relationships (`ListRelationshipsAsync`)
- ✅ Refresh model (`RefreshAsync`)

**✅ Implemented (Excel COM API - Write Operations via TOM):**
- ✅ Create DAX measure (`CreateMeasureAsync`)
- ✅ Update DAX measure (`UpdateMeasureAsync`)
- ✅ Delete DAX measure (`DeleteMeasureAsync`)
- ✅ Create relationship (`CreateRelationshipAsync`)
- ✅ Update relationship (`UpdateRelationshipAsync`)
- ✅ Delete relationship (`DeleteRelationshipAsync`)

**❌ Future (TOM API - Advanced Features):**
- ❌ Calculated columns
- ❌ Hierarchies and perspectives
- ❌ Row-level security (RLS)
- ❌ Advanced DAX validation

**Code Location:** `src/ExcelMcp.Core/Commands/DataModel/`

## User Scenarios & Testing

### User Story 1 - Inspect Data Model Structure (Priority: P1) 🎯 MVP

As a developer, I need to list all tables, columns, and measures in the Data Model to understand the existing BI schema.

**Why this priority**: Foundation for all Data Model work - must see what exists before modifying.

**Independent Test**: Open workbook with Data Model, list tables and measures, verify output matches Excel Power Pivot UI.

**Acceptance Scenarios**:

1. **Given** a workbook with 3 tables in Data Model, **When** I list tables, **Then** I see all 3 with row counts and source information
2. **Given** a table "Sales" with 10 columns, **When** I list columns, **Then** I see all 10 with data types
3. **Given** 5 DAX measures in model, **When** I list measures, **Then** I see all with formulas and parent tables

---

### User Story 2 - Export and Version Control DAX Measures (Priority: P1) 🎯 MVP

As a developer, I need to export DAX measure formulas to files for version control and documentation.

**Why this priority**: Critical for version control, code review, and BI deployment pipelines.

**Independent Test**: Export measure to .dax file, verify file contains formula and metadata, commit to git.

**Acceptance Scenarios**:

1. **Given** a measure "Total Sales", **When** I export to file, **Then** file contains DAX formula with metadata header
2. **Given** 10 measures, **When** I export all, **Then** 10 files created in measures/ directory
3. **Given** exported measures, **When** I diff in git, **Then** formula changes are clearly visible

---

### User Story 3 - Create and Update DAX Measures (Priority: P2)

As a developer, I need to create new DAX measures and update existing ones programmatically.

**Why this priority**: Enables automated measure deployment and DAX refactoring.

**Independent Test**: Create measure, verify in Excel, update formula, verify change persists.

**Acceptance Scenarios**:

1. **Given** a table "Sales", **When** I create measure "Total Revenue = SUM(Sales[Amount])", **Then** measure appears in Power Pivot
2. **Given** existing measure, **When** I update formula, **Then** new formula persists after save/reopen
3. **Given** invalid DAX syntax, **When** I try to create measure, **Then** I get clear validation error

---

### User Story 4 - Manage Table Relationships (Priority: P2)

As a developer, I need to create and modify relationships between Data Model tables.

**Why this priority**: Essential for building proper star schema and data models.

**Independent Test**: Create relationship, verify in diagram view, delete relationship, verify removed.

**Acceptance Scenarios**:

1. **Given** tables "Sales" and "Products", **When** I create relationship on ProductID, **Then** relationship appears in diagram
2. **Given** inactive relationship, **When** I activate it, **Then** it becomes active in model
3. **Given** circular relationship dependency, **When** I try to create, **Then** I get clear error message

---

### User Story 5 - Refresh Data Model (Priority: P3)

As a developer, I need to refresh Data Model tables to update data from source queries.

**Why this priority**: Nice-to-have for data pipeline automation.

**Independent Test**: Modify source data, refresh model, verify data updated in tables.

**Acceptance Scenarios**:

1. **Given** stale Data Model data, **When** I refresh model, **Then** all tables update from sources
2. **Given** source connection timeout, **When** I refresh, **Then** I get timeout error with retry guidance
3. **Given** large datasets (millions of rows), **When** I refresh with 5-min timeout, **Then** operation completes

---

### Edge Cases

- **Large models**: What happens with models containing 50+ measures?
  - ✅ List operations handle large collections efficiently
- **DAX syntax errors**: How are formula errors detected?
  - ✅ TOM API validates on create/update, Excel shows error on refresh
- **Circular relationships**: What happens when creating invalid relationships?
  - ✅ Excel COM API returns error, operation fails safely
- **Measure dependencies**: Can I delete a measure referenced by others?
  - ⚠️ No validation yet - creates orphaned references
- **Concurrent modifications**: What if two processes modify same measure?
  - ✅ Batch API ensures exclusive access during operation

## Requirements

### Functional Requirements

- **FR-001**: System MUST list all tables in Data Model with row counts and source information
- **FR-002**: System MUST list all DAX measures with formulas and parent tables
- **FR-003**: System MUST export DAX measures to .dax files with metadata headers
- **FR-004**: System MUST create new DAX measures with validation via TOM API
- **FR-005**: System MUST update existing DAX measures preserving metadata
- **FR-006**: System MUST delete measures with confirmation
- **FR-007**: System MUST list all relationships with active/inactive status
- **FR-008**: System MUST create relationships between tables with cardinality and filter direction
- **FR-009**: System MUST refresh Data Model with timeout support (default 2 min, max 5 min)
- **FR-010**: System MUST provide model summary (table count, measure count, relationship count)

### Key Entities

- **Model**: Excel PowerPivot Data Model (workbook.Model)
  - Properties: Tables, Measures, Relationships, RefreshDate
  - Operations: Refresh, Initialize

- **ModelTable**: Table in Data Model
  - Properties: Name, SourceName, RecordCount, Columns
  - Source: Power Query, Connection, Excel Table

- **ModelMeasure**: DAX calculation
  - Properties: Name, Formula, Description, AssociatedTable, FormatInformation
  - Created via: TOM API (AMO.dll)

- **ModelRelationship**: Join between tables
  - Properties: ForeignKeyColumn, PrimaryKeyColumn, Active, CrossFilterDirection
  - Types: One-to-Many, Many-to-One

### Non-Functional Requirements

- **NFR-001**: Data Model operations must complete within timeout (2-5 minutes for refresh)
- **NFR-002**: DAX formulas must be validated before committing to model
- **NFR-003**: Measure export must preserve formatting and metadata for readability
- **NFR-004**: COM object cleanup must be guaranteed (try/finally pattern)
- **NFR-005**: TOM API errors must include measure name and formula context

## Success Criteria

### Measurable Outcomes

1. **Read Operations**: All 9 read methods implemented and tested
   - ✅ **ACHIEVED**: List/View/Export operations complete
2. **Write Operations**: Create/Update/Delete for measures and relationships
   - ✅ **ACHIEVED**: Via TOM API, all operations functional
3. **Performance**: Model refresh completes within 5 minutes (95th percentile)
   - ✅ **ACHIEVED**: Timeout support with configurable limits
4. **Test Coverage**: Integration tests for all table sources (Power Query, Connection, Excel Table)
   - ✅ **ACHIEVED**: Tests cover all scenarios
5. **Documentation**: All DAX measures exportable with metadata headers
   - ✅ **ACHIEVED**: Export includes formula, description, format string

### Qualitative Outcomes

- Developers can version control DAX measures alongside M code
- AI agents can discover and analyze Data Model structure
- BI deployment pipelines can programmatically create measures
- Data model health checks can run in CI/CD

## Technical Context

### Excel COM API Used (Layer 1 - Basic Access)

- `Workbook.Model` - Data Model object
- `Model.ModelTables` - Table collection
- `Model.ModelMeasures` - Measure collection (READ-ONLY for formulas)
- `Model.ModelRelationships` - Relationship collection
- `ModelTable.ModelTableColumns` - Column collection

**Limitation**: Excel COM API is READ-ONLY for most operations. Write operations require TOM API.

### Tabular Object Model (TOM) API (Layer 2 - Full Access)

- **NuGet**: `Microsoft.AnalysisServices.NetCore.retail.amd64`
- **Connection**: `Data Source={excelFile};Provider=MSOLAP`
- **Operations**: Full CRUD for measures, relationships, calculated columns
- **Validation**: DAX syntax validation before commit

**Implementation**: `src/ExcelMcp.Core/Commands/DataModel/DataModelCommands.Write.cs`

### Architecture Patterns

- **Two-Layer Access**: Excel COM for read, TOM for write
- **Batch API**: All operations use IExcelBatch for exclusive access
- **Timeout Support**: Refresh accepts configurable timeout
- **Measure Export Format**:
  ```dax
  -- Measure: Total Sales
  -- Table: Sales
  -- Description: Sum of all sales amounts
  -- Format: Currency
  
  Total Sales := SUM(Sales[Amount])
  ```

### Known Limitations

- **Calculated Columns**: Not yet implemented (requires TOM API enhancements)
- **Hierarchies**: Not yet implemented (future TOM feature)
- **Row-Level Security**: Not yet implemented (advanced TOM feature)
- **DAX IntelliSense**: Not available programmatically
- **Model Size**: Large models (>1GB) may hit memory limits during TOM operations

## Testing Strategy

### Integration Tests

- **Test File**: `tests/ExcelMcp.Core.Tests/Commands/DataModelCommandsTests.cs`
- **Test Approach**: Use test workbooks with known Data Model schemas
- **Coverage**:
  - List tables (empty model, multiple tables)
  - List measures (all measures, filtered by table)
  - Export measure to .dax file
  - Create measure via TOM
  - Update measure formula
  - Delete measure
  - List relationships
  - Create relationship
  - Refresh model with timeout

### Manual Test Scenarios

1. Open workbook with Power Pivot model → List tables → Verify matches UI
2. Export all measures → Verify .dax files created
3. Create new measure via TOM → Verify appears in Power Pivot
4. Update measure formula → Verify change persists
5. Refresh model → Verify data updates from source

## Related Documentation

- **Implementation**: `DATA-MODEL-DAX-FEATURE-SPEC.md` (original spec)
- **TOM API Guide**: https://docs.microsoft.com/analysis-services/tom/
- **DAX Reference**: https://dax.guide
- **Testing**: `testing-strategy.instructions.md`
