# Integration Test Conversion Guide

## Overview

This guide documents the proven pattern for converting integration tests from synchronous API to async batch-of-one pattern.

## Status

**Completed:**
- ✅ `CellCommandsTests.cs` - 11 methods
- ✅ `FileCommandsTests.cs` - 10 methods

**Remaining (13+ files, ~232 methods):**
- `ParameterCommandsTests.cs` - 9 methods
- `SheetCommandsTests.cs` - 10 methods
- `CoreConnectionCommandsTests.cs` - ~15 methods
- `CoreConnectionCommandsExtendedTests.cs` - ~20 methods
- `SetupCommandsTests.cs` - ~3 methods
- `VbaTrustDetectionTests.cs` - ~5 methods
- `PowerQueryWorkflowGuidanceTests.cs` - ~30 methods
- `PowerQueryPrivacyLevelTests.cs` - ~8 methods
- `PowerQueryCommandsTests.cs` - ~25 methods
- `DataModelCommandsTests.cs` - ~15 methods
- `DataModelTomCommandsTests.cs` - ~5 methods
- `IntegrationWorkflowTests.cs` (RoundTrip) - ~10 methods
- `ConnectionWorkflowTests.cs` (RoundTrip) - ~12 methods
- `ScriptCommandsRoundTripTests.cs` (RoundTrip) - ~8 methods

**Total Progress:** 21/~253 methods (8%)

---

## Proven Conversion Pattern

### Step 1: Add Using Statement

```csharp
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Session;  // ADD THIS
using Xunit;
```

### Step 2: Convert Constructor (if using file creation)

```csharp
// BEFORE
var result = _fileCommands.CreateEmpty(_testExcelFile);
if (!result.Success) { throw... }

// AFTER
var result = _fileCommands.CreateEmptyAsync(_testExcelFile).GetAwaiter().GetResult();
if (!result.Success) { throw... }
```

### Step 3: Convert Test Method Signatures

```csharp
// BEFORE
[Fact]
public void TestMethod()

// AFTER
[Fact]
public async Task TestMethod()
```

### Step 4: Convert Method Calls Based on Pattern

#### Pattern A: Read-Only Operations (List, Get, View)

```csharp
// BEFORE
var result = _commands.List(_testExcelFile);

// AFTER
await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
var result = await _commands.ListAsync(batch);
```

#### Pattern B: Write Operations (Create, Set, Update, Delete)

```csharp
// BEFORE
var result = _commands.Create(_testExcelFile, "name", "value");

// AFTER
await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
var result = await _commands.CreateAsync(batch, "name", "value");
await batch.SaveAsync();
```

#### Pattern C: Set-Then-Get (Separate Batch Scopes)

```csharp
// BEFORE
var setResult = _commands.Set(_testExcelFile, "param", "value");
var getResult = _commands.Get(_testExcelFile, "param");

// AFTER
await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
{
    var setResult = await _commands.SetAsync(batch, "param", "value");
    Assert.True(setResult.Success);
    await batch.SaveAsync();
}

await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
{
    var getResult = await _commands.GetAsync(batch, "param");
    Assert.Equal("value", getResult.Value?.ToString());
}
```

#### Pattern D: Non-Existent File Tests

```csharp
// BEFORE
var result = _commands.List("nonexistent.xlsx");
Assert.False(result.Success);

// AFTER
await Assert.ThrowsAsync<FileNotFoundException>(async () =>
{
    await using var batch = await ExcelSession.BeginBatchAsync("nonexistent.xlsx");
});
```

---

## Complete Example: ParameterCommandsTests

Here's how to convert `ParameterCommandsTests.cs`:

### 1. Add Using

```csharp
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Session;  // ADD
using Xunit;
```

### 2. Fix Constructor

```csharp
public CoreParameterCommandsTests()
{
    _parameterCommands = new ParameterCommands();
    _fileCommands = new FileCommands();
    
    _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_ParamTests_{Guid.NewGuid():N}");
    Directory.CreateDirectory(_tempDir);
    
    _testExcelFile = Path.Combine(_tempDir, "TestWorkbook.xlsx");
    
    // BEFORE: var result = _fileCommands.CreateEmpty(_testExcelFile);
    // AFTER:
    var result = _fileCommands.CreateEmptyAsync(_testExcelFile).GetAwaiter().GetResult();
    if (!result.Success)
    {
        throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}");
    }
}
```

### 3. Convert Each Test Method

#### List_WithValidFile_ReturnsSuccess (Read-only)

```csharp
[Fact]
public async Task List_WithValidFile_ReturnsSuccess()  // Changed signature
{
    // Act
    await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
    var result = await _parameterCommands.ListAsync(batch);

    // Assert
    Assert.True(result.Success);
    Assert.NotNull(result.Parameters);
}
```

#### Create_WithValidParameter_ReturnsSuccess (Write)

```csharp
[Fact]
public async Task Create_WithValidParameter_ReturnsSuccess()  // Changed signature
{
    // Act
    await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
    var result = await _parameterCommands.CreateAsync(batch, "TestParam", "Sheet1!A1");
    await batch.SaveAsync();  // IMPORTANT: Save after write

    // Assert
    Assert.True(result.Success);
}
```

#### Create_ThenList_ShowsCreatedParameter (Separate batches)

```csharp
[Fact]
public async Task Create_ThenList_ShowsCreatedParameter()
{
    // Arrange
    string paramName = "IntegrationTestParam";

    // Act - Create in one batch
    await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
    {
        var createResult = await _parameterCommands.CreateAsync(batch, paramName, "Sheet1!B2");
        Assert.True(createResult.Success);
        await batch.SaveAsync();
    }

    // Act - List in separate batch
    await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
    {
        var listResult = await _parameterCommands.ListAsync(batch);
        
        // Assert
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Parameters, p => p.Name == paramName);
    }
}
```

#### Set_ThenGet_ReturnsSetValue (Separate batches)

```csharp
[Fact]
public async Task Set_ThenGet_ReturnsSetValue()
{
    // Arrange
    string paramName = "GetSetParam_" + Guid.NewGuid().ToString("N")[..8];
    string testValue = "Integration Test Value";
    
    // Create parameter first
    await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
    {
        var createResult = await _parameterCommands.CreateAsync(batch, paramName, "Sheet1!D1");
        Assert.True(createResult.Success, $"Failed to create parameter: {createResult.ErrorMessage}");
        await batch.SaveAsync();
    }

    // Set value
    await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
    {
        var setResult = await _parameterCommands.SetAsync(batch, paramName, testValue);
        Assert.True(setResult.Success, $"Failed to set parameter: {setResult.ErrorMessage}");
        await batch.SaveAsync();
    }

    // Get value
    await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
    {
        var getResult = await _parameterCommands.GetAsync(batch, paramName);
        Assert.True(getResult.Success, $"Failed to get parameter: {getResult.ErrorMessage}");
        Assert.Equal(testValue, getResult.Value?.ToString());
    }
}
```

#### List_WithNonExistentFile_ReturnsError

```csharp
[Fact]
public async Task List_WithNonExistentFile_ReturnsError()
{
    // Act & Assert
    await Assert.ThrowsAsync<FileNotFoundException>(async () =>
    {
        await using var batch = await ExcelSession.BeginBatchAsync("nonexistent.xlsx");
    });
}
```

---

## Command Mapping Reference

| Old Sync Method | New Async Method | Needs Batch | Needs Save |
|-----------------|------------------|-------------|------------|
| **CellCommands** ||||
| `GetValue(path, ...)` | `GetValueAsync(batch, ...)` | ✅ | ❌ |
| `SetValue(path, ...)` | `SetValueAsync(batch, ...)` | ✅ | ✅ |
| `GetFormula(path, ...)` | `GetFormulaAsync(batch, ...)` | ✅ | ❌ |
| `SetFormula(path, ...)` | `SetFormulaAsync(batch, ...)` | ✅ | ✅ |
| **ParameterCommands** ||||
| `List(path)` | `ListAsync(batch)` | ✅ | ❌ |
| `Get(path, name)` | `GetAsync(batch, name)` | ✅ | ❌ |
| `Set(path, name, value)` | `SetAsync(batch, name, value)` | ✅ | ✅ |
| `Create(path, name, ref)` | `CreateAsync(batch, name, ref)` | ✅ | ✅ |
| `Delete(path, name)` | `DeleteAsync(batch, name)` | ✅ | ✅ |
| **SheetCommands** ||||
| `List(path)` | `ListAsync(batch)` | ✅ | ❌ |
| `Read(path, sheet, range)` | `ReadAsync(batch, sheet, range)` | ✅ | ❌ |
| `Write(path, sheet, csvFile)` | `WriteAsync(batch, sheet, csvFile)` | ✅ | ✅ |
| `Create(path, name)` | `CreateAsync(batch, name)` | ✅ | ✅ |
| `Rename(path, old, new)` | `RenameAsync(batch, old, new)` | ✅ | ✅ |
| `Delete(path, name)` | `DeleteAsync(batch, name)` | ✅ | ✅ |
| `Clear(path, sheet, range)` | `ClearAsync(batch, sheet, range)` | ✅ | ✅ |
| `Copy(path, src, dest)` | `CopyAsync(batch, src, dest)` | ✅ | ✅ |
| `Append(path, sheet, csvFile)` | `AppendAsync(batch, sheet, csvFile)` | ✅ | ✅ |
| **FileCommands** ||||
| `CreateEmpty(path)` | `CreateEmptyAsync(path)` | ❌ | N/A |

---

## Common Mistakes to Avoid

### ❌ Mistake 1: Using batch parameter for CreateEmptyAsync

```csharp
// WRONG
await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
var result = await _fileCommands.CreateEmptyAsync(batch, _testExcelFile);

// RIGHT
var result = await _fileCommands.CreateEmptyAsync(_testExcelFile);
```

### ❌ Mistake 2: Forgetting SaveAsync after write operations

```csharp
// WRONG - Changes not saved!
await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
var result = await _commands.SetAsync(batch, "param", "value");

// RIGHT
await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
var result = await _commands.SetAsync(batch, "param", "value");
await batch.SaveAsync();
```

### ❌ Mistake 3: Reusing batch across set-then-get

```csharp
// WRONG - Same batch sees uncommitted changes
await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
await _commands.SetAsync(batch, "param", "value");
await batch.SaveAsync();
var result = await _commands.GetAsync(batch, "param");  // May fail!

// RIGHT - Separate batches
await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
{
    await _commands.SetAsync(batch, "param", "value");
    await batch.SaveAsync();
}
await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
{
    var result = await _commands.GetAsync(batch, "param");
}
```

### ❌ Mistake 4: Not converting constructor file creation

```csharp
// WRONG - Old sync API
var result = _fileCommands.CreateEmpty(_testExcelFile);

// RIGHT - New async API with sync wrapper
var result = _fileCommands.CreateEmptyAsync(_testExcelFile).GetAwaiter().GetResult();
```

---

## Build Verification

After converting each file:

```powershell
# Build to check for errors
dotnet build tests/ExcelMcp.Core.Tests --no-restore

# Count remaining errors
dotnet build tests/ExcelMcp.Core.Tests --no-restore 2>&1 | Select-String -Pattern "(\d+) Error\(s\)"
```

**Expected Error Reduction:**
- ParameterCommandsTests: ~15-20 errors
- SheetCommandsTests: ~18-25 errors
- Each connection test file: ~30-50 errors
- PowerQuery files: ~50-80 errors each

---

## Workflow

1. Pick a test file from the Remaining list
2. Add `using Sbroenne.ExcelMcp.Core.Session;`
3. Fix constructor if it uses `CreateEmpty`
4. Convert each test method:
   - Change `void` → `async Task`
   - Wrap calls in batch-of-one pattern
   - Add `SaveAsync()` for write operations
   - Use separate batches for set-then-get
5. Build and verify error count decreased
6. Commit when file is complete

---

## Automation Script

A PowerShell script is available at `scripts/Convert-TestsToBatchPattern.ps1` for partial automation, but manual review is required for:
- Batch scope separation (set-then-get)
- SaveAsync placement
- Non-existent file test conversions

---

## Progress Tracking

Update this section as files are completed:

```
✅ CellCommandsTests.cs (11 methods)
✅ FileCommandsTests.cs (10 methods)
⬜ ParameterCommandsTests.cs (9 methods)
⬜ SheetCommandsTests.cs (10 methods)
... (see Remaining list above)
```

**Current:** 21/253 methods (8%)  
**Target:** 253/253 methods (100%)

---

## Questions?

- See `CellCommandsTests.cs` for complete working example
- See `FileCommandsTests.cs` for CreateEmptyAsync pattern
- Check commit history for proven patterns
- All Core commands are already async - just need batch parameter
