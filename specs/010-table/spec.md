# Feature Specification: Excel Table Management

**Feature Branch**: `010-table`  
**Created**: 2024-01-10  
**Status**: ✅ **IMPLEMENTED**  
**Last Updated**: 2025-11-10

## Implementation Status

**✅ FULLY IMPLEMENTED** - All Excel Table (ListObject) operations functional.

**✅ Implemented:**
- ✅ Lifecycle (List, Get, Create, Rename, Delete, Resize)
- ✅ Style management (SetStyleAsync)
- ✅ Totals row (ToggleTotalsAsync, SetColumnTotalAsync)
- ✅ Append data (AppendAsync)
- ✅ Add to Data Model (AddToDataModelAsync)
- ✅ Filtering (ApplyFilterAsync, ApplyFilterValuesAsync, ClearFiltersAsync, GetFiltersAsync)
- ✅ Column operations (AddColumnAsync, RemoveColumnAsync, RenameColumnAsync)
- ✅ Structured references (GetStructuredReferenceAsync)
- ✅ Sorting (SortAsync single column, SortMultiAsync multiple columns)
- ✅ Number formatting (GetColumnNumberFormatAsync, SetColumnNumberFormatAsync)

**Code Location:** `src/ExcelMcp.Core/Commands/Table/`

## User Scenarios

### User Story 1 - Create and Manage Tables (Priority: P1) 🎯 MVP

As a developer, I need to create Excel Tables from ranges for structured data.

**Acceptance Scenarios**:
1. **Given** range A1:D10 with headers, **When** I create table "Sales", **Then** table created with AutoFilter
2. **Given** table "Sales", **When** I rename to "Revenue", **Then** name changes
3. **Given** table, **When** I resize to A1:F20, **Then** table expands
4. **Given** table, **When** I delete, **Then** table removed, data remains

### User Story 2 - Apply Table Styles (Priority: P1) 🎯 MVP

As a developer, I need to apply built-in table styles for professional appearance.

**Acceptance Scenarios**:
1. **Given** table, **When** I apply "TableStyleMedium2", **Then** table styled with blue theme
2. **Given** table, **When** I apply "TableStyleLight1", **Then** table styled with light theme

### User Story 3 - Manage Totals Row (Priority: P2)

As a developer, I need to add totals row with aggregation functions.

**Acceptance Scenarios**:
1. **Given** table, **When** I toggle totals on, **Then** totals row appears
2. **Given** column "Sales", **When** I set total to SUM, **Then** total calculates
3. **Given** column "Count", **When** I set total to COUNT, **Then** count displays

### User Story 4 - Filter and Sort Tables (Priority: P2)

As a developer, I need to filter and sort table data programmatically.

**Acceptance Scenarios**:
1. **Given** status column, **When** I filter by ["Open", "Pending"], **Then** only matching rows visible
2. **Given** table, **When** I sort by sales descending, **Then** rows reordered

### User Story 5 - Append Data to Tables (Priority: P2)

As a developer, I need to add new rows to tables dynamically.

**Acceptance Scenarios**:
1. **Given** table with 100 rows, **When** I append CSV data, **Then** new rows added, table auto-expands

## Requirements

### Functional Requirements
- **FR-001**: Create table from range with optional headers
- **FR-002**: List all tables with metadata
- **FR-003**: Get table details (name, range, style, row count)
- **FR-004**: Rename table
- **FR-005**: Delete table (preserve data)
- **FR-006**: Resize table range
- **FR-007**: Apply built-in table styles
- **FR-008**: Toggle totals row on/off
- **FR-009**: Set column total function (SUM, AVG, COUNT, etc.)
- **FR-010**: Append rows via CSV data
- **FR-011**: Add table to Data Model
- **FR-012**: Apply column filters (criteria or value list)
- **FR-013**: Clear all filters
- **FR-014**: Get filter status
- **FR-015**: Add/Remove/Rename columns
- **FR-016**: Get structured reference formula (e.g., "Sales[Amount]")
- **FR-017**: Sort single column
- **FR-018**: Sort multiple columns with priority
- **FR-019**: Get/Set column number formats

### Non-Functional Requirements
- **NFR-001**: Table operations complete within 2 minutes
- **NFR-002**: Table auto-expands when appending data
- **NFR-003**: Structured references update automatically on rename

## Success Criteria
- ✅ All 19 table methods implemented
- ✅ Integration with Data Model confirmed
- ✅ CSV append functionality tested

## Technical Context

### Excel COM API
- `Worksheet.ListObjects` - Table collection
- `ListObjects.Add()` - Create table
- `ListObject.TableStyle` - Built-in style name
- `ListObject.ShowTotals` - Toggle totals row
- `ListObject.Resize()` - Change range
- `ListObject.DataBodyRange` - Data rows (excluding headers/totals)
- `ListObject.ListRows.Add()` - Append row
- `ListObject.AutoFilter.Filters()` - Column filters

### Architecture
- Partial classes: TableCommands split by operation type
- Batch API for exclusive access
- CSV parsing for append operations
- Structured reference formula generation

## Related Documentation
- **Original Spec**: `TABLE-API-SPECIFICATION.md`
- **Testing**: `testing-strategy.instructions.md`
