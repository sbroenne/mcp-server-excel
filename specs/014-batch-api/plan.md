# Implementation Plan: Batch API for Multi-Operation Workflows

**Feature**: Batch API  
**Branch**: `014-batch-api`  
**Status**: ✅ **IMPLEMENTED**  
**Last Updated**: 2025-11-10

## Implementation Status

### ✅ ALL COMPLETE
- ✅ ExcelSession.BeginBatchAsync
- ✅ IExcelBatch interface
- ✅ ExcelBatch implementation
- ✅ STA threading with OLE message filter
- ✅ Explicit save pattern
- ✅ Timeout support
- ✅ COM cleanup

## Architecture

### Component Structure
```
src/ExcelMcp.ComInterop/Session/
├── IExcelBatch.cs                  # Public interface
├── ExcelBatch.cs                   # Implementation
├── ExcelSession.cs                 # Static entry point
└── ExcelContext.cs                 # Operation context
```

### Supporting Infrastructure
```
src/ExcelMcp.ComInterop/
├── OleMessageFilter.cs             # Cross-thread COM handling
├── IOleMessageFilter.cs            # COM interface
├── ComUtilities.cs                 # COM cleanup helpers
└── FileAccessValidator.cs          # Path validation
```

## Technology Stack
- .NET 9.0, Windows-only
- Excel COM Interop (dynamic types)
- STA threading via OLE message filter
- IAsyncDisposable pattern
- Timeout via CancellationTokenSource

## Key Design Decisions

### Decision 1: Explicit Save Pattern
**Pattern**:
```csharp
await using var batch = await ExcelSession.BeginBatchAsync(filePath);
await batch.ExecuteAsync(...);
await batch.SaveAsync();  // EXPLICIT save required
// If SaveAsync not called, changes discarded on dispose
```

**Why**: Try-before-commit workflows, prevents accidental persistence

### Decision 2: IAsyncDisposable
**Implementation**: `await using` ensures cleanup
**Benefits**: Automatic COM cleanup, file lock release, GC collection

### Decision 3: STA Threading with OLE Filter
**Challenge**: Excel COM requires STA thread
**Solution**: IOleMessageFilter handles cross-thread marshalling
**Registration**: `CoRegisterMessageFilter()` in ExcelSession

### Decision 4: Timeout Architecture
**Levels**:
- Default: 2 minutes (most operations)
- Extended: 5 minutes (refresh, data model)
- Per-operation override: ExecuteAsync accepts timeout parameter

**Implementation**: CancellationTokenSource with timeout

### Decision 5: Exclusive File Access
**Enforcement**: Track open files in ExcelSession
**Why**: Excel COM cannot handle concurrent workbook access

### Decision 6: COM Cleanup Guarantee
**Pattern**:
```csharp
dynamic? comObject = null;
try {
    comObject = workbook.SomeProperty;
    // Use comObject...
} finally {
    ComUtilities.Release(ref comObject);
}
```

**OLE Filter**: Handles "RPC_E_SERVERCALL_RETRYLATER" busy messages

## Testing Strategy

### Integration Tests
- **Location**: `tests/ExcelMcp.ComInterop.Tests/Session/`
- **Traits**: `[Trait("RunType", "OnDemand")]` (slow tests)
- **Coverage**:
  - Batch creation and disposal
  - Execute with save vs discard
  - Timeout scenarios
  - COM leak detection
  - Performance benchmarks

### Test Execution
```powershell
# Run batch/session tests (SLOW - only when changing session code)
dotnet test tests/ExcelMcp.ComInterop.Tests/ExcelMcp.ComInterop.Tests.csproj --filter "RunType=OnDemand"
```

### Pre-Commit Hook
Script `scripts/check-com-leaks.ps1` scans for COM object patterns without cleanup

## Performance Characteristics

**Measured** (from integration tests):
- Batch (10 operations): ~2-3 seconds
- Individual (10 operations): ~10-15 seconds
- **Speedup**: 75-85%

**Why Batch is Faster**:
1. Single Excel.Application instance
2. Single Workbook.Open()
3. Reduced COM marshalling overhead
4. No file system contention

## Security Considerations

- **File Access**: Validated via FileAccessValidator before opening
- **Path Traversal**: Prevented by absolute path requirement
- **COM Security**: Excel COM runs in user context (no elevation)
- **File Locks**: Released on dispose (no orphaned locks)

## Deployment Considerations

### Runtime Requirements
- Excel installed (COM interop)
- .NET 9.0 runtime
- Windows OS (COM + STA threading)

### CI/CD
- Integration tests require Excel
- Tests run on Azure self-hosted runner
- Cannot run in GitHub Actions hosted (no Excel)

## Known Limitations

- **Concurrent Access**: Cannot have multiple batches on same file
- **Excel UI**: Operations may briefly show Excel window (hidden but visible)
- **COM Exceptions**: Excel "busy" errors handled by OLE filter, but some scenarios may fail
- **File Locks**: If process crashes, Excel may leave file locked (manual recovery)

## Migration Notes

### Breaking Changes
None - Batch API is foundational feature in v1.0.0

### Evolution from WithExcel()
**Before** (deprecated): `ExcelHelper.WithExcel(filePath, save, (excel, workbook) => { ... })`
**After** (current): `await using var batch = await ExcelSession.BeginBatchAsync(filePath); await batch.ExecuteAsync(...); await batch.SaveAsync();`

**Why Changed**: Explicit save control, better async support, cleaner disposal pattern

## Related Documentation

- **Spec**: `014-batch-api/spec.md`
- **Testing**: `.github/instructions/testing-strategy.instructions.md`
- **Excel COM**: `.github/instructions/excel-com-interop.instructions.md`
- **Architecture**: `.github/instructions/architecture-patterns.instructions.md`
- **Timeout Guide**: `docs/TIMEOUT-IMPLEMENTATION-GUIDE.md`
- **Critical Rules**: `.github/instructions/critical-rules.instructions.md` (Rule 3 - Session Cleanup Tests)
