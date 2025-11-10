# Feature Specification: Power Query M Code Management

**Feature Branch**: `006-powerquery`  
**Created**: 2024-01-10  
**Status**: ✅ **IMPLEMENTED**  
**Last Updated**: 2025-11-10

## Implementation Status

**✅ FULLY IMPLEMENTED** - All Power Query operations functional.

**✅ Implemented:**
- ✅ List queries
- ✅ View M code
- ✅ Import from .pq files
- ✅ Export to .pq files
- ✅ Update M code
- ✅ Refresh queries
- ✅ Delete queries
- ✅ Load destination management (worksheet, data-model, both, connection-only)
- ✅ Get load configuration
- ✅ List Excel data sources
- ✅ Evaluate M code

**Code Location:** `src/ExcelMcp.Core/Commands/PowerQuery/`

## User Scenarios

### User Story 1 - Version Control M Code (Priority: P1) 🎯 MVP

As a developer, I need to export Power Query M code to files for version control.

**Acceptance Scenarios**:
1. **Given** query "SalesData", **When** I export to file, **Then** .pq file contains M code
2. **Given** 5 queries, **When** I export all, **Then** 5 .pq files created
3. **Given** exported M code, **When** I diff in git, **Then** changes are visible

### User Story 2 - Import and Update Queries (Priority: P1) 🎯 MVP

As a developer, I need to import M code from files and update existing queries.

**Acceptance Scenarios**:
1. **Given** .pq file, **When** I import, **Then** query appears in workbook
2. **Given** existing query, **When** I update from file, **Then** M code replaced
3. **Given** updated query, **When** I refresh, **Then** data reloads from source

### User Story 3 - Manage Load Destinations (Priority: P2)

As a developer, I need to control where query data loads (worksheet, Data Model, both, or connection-only).

**Acceptance Scenarios**:
1. **Given** query, **When** I set loadDestination='worksheet', **Then** data loads to table
2. **Given** query, **When** I set loadDestination='data-model', **Then** data loads to Power Pivot
3. **Given** query, **When** I set loadDestination='connection-only', **Then** no data loads

### User Story 4 - Refresh Queries (Priority: P2)

As a developer, I need to refresh Power Query data from sources.

**Acceptance Scenarios**:
1. **Given** stale query, **When** I refresh, **Then** data updates from source
2. **Given** slow data source, **When** I refresh with 5-min timeout, **Then** operation completes

## Requirements

### Functional Requirements
- **FR-001**: List all Power Queries with names and load destinations
- **FR-002**: View M code for any query
- **FR-003**: Import M code from .pq files
- **FR-004**: Export M code to .pq files
- **FR-005**: Update existing query M code
- **FR-006**: Refresh query data
- **FR-007**: Delete queries
- **FR-008**: Set load destination (worksheet, data-model, both, connection-only)
- **FR-009**: Get current load configuration
- **FR-010**: List Excel sources (ranges, tables) for M code generation

### Non-Functional Requirements
- **NFR-001**: Refresh operations support 5-minute timeout
- **NFR-002**: M code preserved exactly (no reformatting)
- **NFR-003**: Import validates M syntax before committing

## Success Criteria
- ✅ All 11 Power Query methods implemented
- ✅ M code version control workflows functional
- ✅ Integration with Data Model confirmed

## Technical Context

### Excel COM API
- `Workbook.Queries` - Query collection
- `WorkbookQuery.Formula` - M code (read-only for write)
- `QueryTables.Add()` - Load query to worksheet
- Connection objects for metadata

### Architecture
- Partial classes for query operations
- QueryTable pattern for data loading
- Refresh(false) for synchronous persistence

## Related Documentation
- **Original Spec**: `POWERQUERY-FUTURE-STATE-SPEC.md`
- **Testing**: `testing-strategy.instructions.md`
