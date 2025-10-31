# Async Pattern Refactoring - Implementation Guide

## Completed ✅

### 1. Dual Interface in ComInterop Layer
**File**: `src/ExcelMcp.ComInterop/Session/IExcelBatch.cs`

```csharp
// For synchronous COM operations (most operations)
Task<T> Execute<T>(
    Func<ExcelContext, CancellationToken, T> operation,
    CancellationToken cancellationToken = default);

// For genuinely async operations (file I/O only)
Task<T> ExecuteAsync<T>(
    Func<ExcelContext, CancellationToken, Task<T>> operation,
    CancellationToken cancellationToken = default);
```

**Why This Works:**
- `Execute()`: Lambda returns `T` directly - for synchronous COM calls
- `ExecuteAsync()`: Lambda returns `Task<T>` - for async file I/O
- Both marshal to STA thread correctly
- Type system enforces correct usage

### 2. ExcelSession Dual Methods
**File**: `src/ExcelMcp.ComInterop/Session/ExcelSession.cs`

```csharp
// Synchronous COM operations
Task<T> CreateNew<T>(string filePath, bool isMacroEnabled, 
    Func<ExcelContext, CancellationToken, T> operation, ...);

// Async I/O operations
Task<T> CreateNewAsync<T>(string filePath, bool isMacroEnabled,
    Func<ExcelContext, CancellationToken, Task<T>> operation, ...);
```

### 3. Test Files Updated
- All test files now use `Execute()` for pure COM operations
- Removed all `ValueTask.FromResult()` wrapping
- Tests are cleaner and more readable

### 4. Build Status
- ✅ All source projects build successfully
- ✅ Unit tests pass (expected platform failures on non-Windows)

## Remaining Work 🔧

### 1. Update Core Commands (High Priority)

**Current Pattern (WRONG):**
```csharp
#pragma warning disable CS1998  // ❌ Code smell
public async Task<TableListResult> ListAsync(IExcelBatch batch)
{
    return await batch.ExecuteAsync(async (ctx, ct) => {  // ❌ Unnecessary async
        // Only synchronous COM calls
        dynamic tables = ctx.Book.ListObjects;
        return result;
    });
}
```

**New Pattern (CORRECT):**
```csharp
// No pragma needed! ✅
public Task<TableListResult> List(IExcelBatch batch)  // ✅ Removed "Async" suffix
{
    return batch.Execute((ctx, ct) => {  // ✅ Synchronous lambda
        // Synchronous COM calls
        dynamic tables = ctx.Book.ListObjects;
        return result;
    });
}
```

**Exception - Keep Async for File I/O:**
```csharp
// Keep "Async" suffix and async lambda ✅
public Task<OperationResult> ExportAsync(IExcelBatch batch, string file)
{
    return batch.ExecuteAsync(async (ctx, ct) => {
        dynamic query = ctx.Book.Queries.Item(name);
        string mCode = query.Formula;
        await File.WriteAllTextAsync(file, mCode, ct);  // Genuine async
        return new OperationResult { Success = true };
    });
}
```

### 2. Remove Method "Async" Suffixes

**Files to Update (~30 command files):**
```
src/ExcelMcp.Core/Commands/
├── ConnectionCommands.cs
├── DataModel/DataModelCommands.*.cs
├── FileCommands.cs
├── ParameterCommands.cs
├── PivotTable/PivotTableCommands.*.cs
├── PowerQueryCommands.cs
├── Range/RangeCommands.*.cs
├── ScriptCommands.cs
├── SetupCommands.cs
├── SheetCommands.cs
└── Table/TableCommands.*.cs
```

**Method Renaming Examples:**
- `ListAsync()` → `List()`
- `CreateAsync()` → `Create()`
- `UpdateAsync()` → `Update()`
- `DeleteAsync()` → `Delete()`
- **KEEP**: `ExportAsync()`, `ImportAsync()` (do file I/O)

### 3. Remove ALL `#pragma warning disable CS1998`

**Search and verify:**
```bash
grep -r "pragma warning disable CS1998" src/ExcelMcp.Core/Commands
```

After updating to synchronous Execute(), ALL pragmas should be removable.

### 4. Update CLI Commands

**Current Pattern:**
```csharp
var task = Task.Run(async () => {
    await using var batch = await ExcelSession.BeginBatchAsync(filePath);
    var result = await _commands.ListAsync(batch);  // Old
    return result;
});
```

**New Pattern:**
```csharp
var task = Task.Run(async () => {
    await using var batch = await ExcelSession.BeginBatchAsync(filePath);
    var result = await _commands.List(batch);  // New (method renamed)
    return result;
});
```

### 5. Update MCP Server Tools

**Current:**
```csharp
await using var batch = await ExcelSession.BeginBatchAsync(excelPath);
var result = await _commands.ListAsync(batch);  // Old
```

**New:**
```csharp
await using var batch = await ExcelSession.BeginBatchAsync(excelPath);
var result = await _commands.List(batch);  // New
```

### 6. Update Interface Definitions

All `I*Commands` interfaces need method names updated:
- `ITableCommands.ListAsync()` → `ITableCommands.List()`
- etc.

## Implementation Strategy

### Phase 1: Core Commands (One File at a Time)
```bash
# For each command file:
# 1. Change batch.ExecuteAsync( to batch.Execute(
# 2. Remove async keyword from lambda
# 3. Remove #pragma warning disable CS1998
# 4. Remove "Async" suffix from method name (if no file I/O)
# 5. Update interface
# 6. Build and verify
```

### Phase 2: CLI Layer
```bash
# Update all CLI commands to call renamed methods
# Build and verify
```

### Phase 3: MCP Server Layer
```bash
# Update all MCP tools to call renamed methods
# Build and verify
```

### Phase 4: Final Verification
```bash
# Remove documentation about ASYNC-PATTERN-RATIONALE.md (now obsolete)
# Run full test suite
# Update contributing guidelines
```

## Automated Migration Script

```bash
#!/bin/bash
# migrate_command_file.sh - Migrate a single command file

FILE=$1

if [ ! -f "$FILE" ]; then
    echo "File not found: $FILE"
    exit 1
fi

# 1. Change ExecuteAsync to Execute for sync operations
sed -i 's/batch\.ExecuteAsync(async (ctx, ct)/batch.Execute((ctx, ct)/g' "$FILE"
sed -i 's/batch\.ExecuteAsync<\([^>]*\)>(async (ctx, ct)/batch.Execute<\1>((ctx, ct)/g' "$FILE"

# 2. Remove pragma warnings
sed -i '/^#pragma warning disable CS1998/d' "$FILE"
sed -i '/^#pragma warning restore CS1998/d' "$FILE"

# 3. Method renaming (only for methods WITHOUT file I/O)
# This requires manual review per file

echo "File $FILE migrated - review and test before committing"
```

## Testing Strategy

**After each file migration:**
```bash
# 1. Build the specific project
dotnet build src/ExcelMcp.Core

# 2. Run related tests
dotnet test --filter "ClassName~CommandsTests"

# 3. Verify no warnings
dotnet build --no-incremental | grep "warning CS1998"  # Should be empty
```

## Success Criteria

- [ ] Zero `#pragma warning disable CS1998` in src/
- [ ] All synchronous COM methods use `Execute()`
- [ ] Only file I/O methods use `ExecuteAsync()`  
- [ ] Method names reflect sync vs async nature
- [ ] All tests pass
- [ ] Build produces zero warnings

## Benefits Achieved

1. **Code Clarity**: Synchronous operations look synchronous
2. **Type Safety**: Compiler enforces sync vs async
3. **No Pragmas**: No code smells or warning suppressions
4. **Better Naming**: Method names indicate true behavior
5. **Maintainability**: Clear which operations do I/O
