# Feature Specification: QueryTable Support

**Feature Branch**: `007-querytable`  
**Created**: 2024-01-10  
**Status**: ✅ **IMPLEMENTED**  
**Last Updated**: 2025-11-10

## Implementation Status

**✅ FULLY IMPLEMENTED** - All QueryTable operations functional.

**✅ Implemented:**
- ✅ List QueryTables
- ✅ Get QueryTable info
- ✅ Refresh QueryTable
- ✅ Delete QueryTable
- ✅ Update connection properties

**Code Location:** `src/ExcelMcp.Core/Commands/QueryTable/`

## User Scenarios

### User Story 1 - List and Inspect QueryTables (Priority: P1) 🎯 MVP

As a developer, I need to list all QueryTables and view their configuration.

**Acceptance Scenarios**:
1. **Given** workbook with 3 QueryTables, **When** I list, **Then** I see all 3 with connection info
2. **Given** QueryTable "SalesData", **When** I get info, **Then** I see connection string, SQL, refresh settings

### User Story 2 - Refresh QueryTable Data (Priority: P1) 🎯 MVP

As a developer, I need to refresh QueryTable data from external sources.

**Acceptance Scenarios**:
1. **Given** stale QueryTable, **When** I refresh, **Then** data updates from source
2. **Given** slow source, **When** I refresh with timeout, **Then** operation completes or errors gracefully

### User Story 3 - Manage QueryTable Lifecycle (Priority: P2)

As a developer, I need to delete QueryTables or update their connections.

**Acceptance Scenarios**:
1. **Given** QueryTable, **When** I delete, **Then** QueryTable removed, data remains
2. **Given** QueryTable, **When** I update connection string, **Then** next refresh uses new source

## Requirements

### Functional Requirements
- **FR-001**: List all QueryTables in workbook
- **FR-002**: Get QueryTable info (name, connection, command text, refresh settings)
- **FR-003**: Refresh QueryTable with timeout support
- **FR-004**: Delete QueryTable
- **FR-005**: Update connection properties

### Non-Functional Requirements
- **NFR-001**: Refresh supports 5-minute timeout
- **NFR-002**: COM object cleanup guaranteed

## Success Criteria
- ✅ All 5 QueryTable methods implemented
- ✅ Integration with Power Query confirmed

## Technical Context

### Excel COM API
- `Worksheet.QueryTables` - QueryTable collection
- `QueryTable.Refresh(false)` - Synchronous refresh
- `QueryTable.Delete()` - Remove QueryTable

### Architecture
- QueryTable operations integrated with PowerQuery commands
- Refresh uses timeout support

## Related Documentation
- **Original Spec**: `QUERYTABLE-SUPPORT-SPECIFICATION.md`
- **Power Query**: `006-powerquery/spec.md`
