# Feature Specification: Batch API for Multi-Operation Workflows

**Feature Branch**: `014-batch-api`  
**Created**: 2024-01-10  
**Status**: ✅ **IMPLEMENTED**  
**Last Updated**: 2025-11-10

## Implementation Status

**✅ FULLY IMPLEMENTED** - Complete batch session management.

**✅ Implemented:**
- ✅ BeginBatchAsync - Create batch session
- ✅ ExecuteAsync - Execute operations in session
- ✅ SaveAsync - Explicit save with timeout
- ✅ Dispose - Auto-close without save
- ✅ STA threading pattern
- ✅ Exclusive file access
- ✅ Timeout support (default 2 min, configurable up to 5 min)
- ✅ COM object cleanup

**Code Location:** `src/ExcelMcp.ComInterop/Session/`

## User Scenarios

### User Story 1 - Multi-Operation Workflow (Priority: P1) 🎯 MVP

As a developer, I need to perform multiple operations on one workbook without re-opening.

**Acceptance Scenarios**:
1. **Given** 5 operations needed, **When** I use batch mode, **Then** workbook opens once
2. **Given** batch session, **When** I execute operations, **Then** all succeed or fail together
3. **Given** operations complete, **When** I save, **Then** all changes persist

### User Story 2 - Explicit Save Control (Priority: P1) 🎯 MVP

As a developer, I need control over when changes are saved to disk.

**Acceptance Scenarios**:
1. **Given** batch operations complete, **When** I dispose without save, **Then** changes discarded
2. **Given** operations complete, **When** I call SaveAsync, **Then** changes persist
3. **Given** SaveAsync called, **When** I dispose, **Then** no double-save

### User Story 3 - Performance Optimization (Priority: P2)

As a developer, I need batch mode to be 75-90% faster than individual operations.

**Acceptance Scenarios**:
1. **Given** 10 operations, **When** I use batch mode, **Then** 80%+ faster than separate calls
2. **Given** batch session, **When** I execute 100 operations, **Then** single Excel instance used

### User Story 4 - Timeout Support (Priority: P2)

As a developer, I need configurable timeouts for heavy operations.

**Acceptance Scenarios**:
1. **Given** quick operation, **When** I use default 2-min timeout, **Then** completes successfully
2. **Given** data model refresh, **When** I request 5-min timeout, **Then** extended time granted
3. **Given** timeout exceeded, **When** operation runs too long, **Then** TimeoutException with retry guidance

## Requirements

### Functional Requirements
- **FR-001**: Begin batch session on file path
- **FR-002**: Execute operations within exclusive session
- **FR-003**: Save changes explicitly via SaveAsync
- **FR-004**: Auto-dispose without save (discard changes)
- **FR-005**: Support default timeout (2 min) and extended timeout (5 min)
- **FR-006**: Maintain STA threading for Excel COM
- **FR-007**: Ensure exclusive file access (no concurrent batches on same file)
- **FR-008**: Guarantee COM object cleanup on dispose

### Non-Functional Requirements
- **NFR-001**: Batch operations 75-90% faster than individual calls
- **NFR-002**: Timeout exceptions include retry guidance
- **NFR-003**: COM object leaks prevented via OLE message filter
- **NFR-004**: File locks released on dispose

## Success Criteria
- ✅ Batch API implemented with IExcelBatch interface
- ✅ Integration tests confirm performance gains
- ✅ Timeout support for heavy operations
- ✅ COM cleanup guaranteed (0 leaks)

## Technical Context

### Architecture

```
ExcelSession (static)
    └── BeginBatchAsync(filePath) → IExcelBatch
            ├── ExecuteAsync<T>(func) → T
            ├── SaveAsync() → void
            └── DisposeAsync() → close without save
```

### Core Components

**ExcelSession** (`src/ExcelMcp.ComInterop/Session/ExcelSession.cs`):
- Static entry point: `BeginBatchAsync(filePath)`
- Creates ExcelBatch instance
- Manages STA threading

**ExcelBatch** (`src/ExcelMcp.ComInterop/Session/ExcelBatch.cs`):
- Implements IExcelBatch
- Holds Excel.Application and Workbook references
- Exclusive access to file
- COM cleanup on dispose

**IExcelBatch** (`src/ExcelMcp.ComInterop/Session/IExcelBatch.cs`):
- Public interface for batch operations
- `ExecuteAsync<T>` for operations
- `SaveAsync` for persistence
- IAsyncDisposable for cleanup

### STA Threading Pattern

**Challenge**: Excel COM requires STA apartment thread
**Solution**: OLE message filter handles cross-thread calls
**Implementation**: `IOleMessageFilter` registered in ExcelSession

### Key Design Decisions

#### Decision 1: Explicit Save Pattern
**Rationale**: Developers must intentionally save changes
**Default Behavior**: Dispose without save = discard changes
**Why**: Prevents accidental persistence, supports try-before-commit workflows

#### Decision 2: Exclusive File Access
**Implementation**: Only one batch per file at a time
**Why**: Excel COM cannot handle concurrent access

#### Decision 3: Timeout Defaults
**Default**: 2 minutes (most operations)
**Extended**: 5 minutes (refresh, data model, large ranges)
**Why**: Balance responsiveness vs operation needs

#### Decision 4: COM Cleanup Guarantee
**Pattern**: try/finally with ReleaseComObject + GC.Collect()
**OLE Filter**: Handles "busy" messages from Excel
**Why**: Prevents zombie Excel processes

## Testing Strategy

### Integration Tests
- **Tests**: `tests/ExcelMcp.ComInterop.Tests/Session/`
- **Coverage**:
  - Begin batch → Execute operations → SaveAsync
  - Begin batch → Execute operations → Dispose (no save)
  - Timeout scenarios (default, extended, exceeded)
  - COM leak detection (0 leaks after dispose)
  - Performance benchmarks (batch vs individual)

### Test Attributes
- `[Trait("RunType", "OnDemand")]` - Batch tests are slow (~20s each)
- Run explicitly: `dotnet test --filter "RunType=OnDemand"`

### Manual Testing
1. Create batch → Execute 10 operations → SaveAsync → Verify all changes persist
2. Create batch → Execute operations → Dispose → Verify no changes saved
3. Create batch → Long refresh (4 min) → Verify completes with 5-min timeout
4. Create batch → Timeout operation → Verify TimeoutException with guidance

## Performance Metrics

**Measured Performance** (from tests):
- Individual operations (10x): ~10-15 seconds
- Batch operations (10x): ~2-3 seconds
- **Improvement**: 75-85% faster

**Why**:
- Single Excel instance (no open/close overhead)
- Single workbook open
- Reduced COM marshalling

## Related Documentation
- **Testing**: `testing-strategy.instructions.md`
- **Excel COM**: `excel-com-interop.instructions.md`
- **Architecture**: `architecture-patterns.instructions.md`
- **Timeout Guide**: `docs/TIMEOUT-IMPLEMENTATION-GUIDE.md`
