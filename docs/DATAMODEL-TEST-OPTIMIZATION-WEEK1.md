# DataModel Test Optimization - Week 1 Implementation

## Goal
Reduce DataModel test execution from **120s to 55s** (54% improvement) by using template-based test files for READ-ONLY tests.

## Current Situation
- ✅ Template strategy code exists in `DataModelCommandsTests.cs`
- ✅ Template file created on first run: `bin/Debug/net8.0/TestAssets/DataModelTemplate.xlsx`
- ❌ **NO tests actually use the template!** All tests default to fresh Data Model creation

## Implementation Checklist

### Step 1: Identify READ-ONLY vs WRITE Tests

**READ-ONLY Tests (13 total) - Can use template:**
These tests only LIST, VIEW, GET operations - don't modify Data Model.

**File: `DataModelCommandsTests.Tables.cs`**
- [ ] `ListTables_WithRealisticDataModel_ReturnsTablesWithData`
- [ ] `ListTables_WithValidFile_ReturnsSuccessResult`

**File: `DataModelCommandsTests.Discovery.cs`**
- [ ] `ListTableColumns_WithValidTable_ReturnsColumns`
- [ ] `ViewTable_WithValidTable_ReturnsCompleteInfo`
- [ ] `ViewTable_WithTableHavingMeasures_CountsMeasuresCorrectly`
- [ ] `GetModelInfo_WithRealisticDataModel_ReturnsAccurateStatistics`

**File: `DataModelCommandsTests.Measures.cs`**
- [ ] `ListMeasures_WithRealisticDataModel_ReturnsMeasuresWithFormulas`
- [ ] `ViewMeasure_WithRealisticDataModel_ReturnsValidDAXFormula`

**File: `DataModelCommandsTests.Relationships.cs`**
- [ ] `ListRelationships_WithValidFile_ReturnsSuccessResult`
- [ ] `ListRelationships_WithRealisticDataModel_ReturnsRelationshipsWithTables`
- [ ] `ViewRelationship_WithValidRelationship_ReturnsCompleteInfo`
- [ ] `GetRelationshipProperties_WithActiveRelationship_ReturnsActiveState`
- [ ] `GetRelationshipProperties_WithInactiveRelationship_ReturnsInactiveState`

**WRITE Tests (3 total) - MUST keep fresh build:**
These tests CREATE, UPDATE, DELETE operations - need writable Data Model.

**File: `DataModelCommandsTests.Measures.cs`**
- [ ] `CreateMeasure_WithValidParameters_CreatesSuccessfully` (WRITE - keep fresh)
- [ ] `CreateMeasure_WithFormatType_CreatesWithFormat` (WRITE - keep fresh)
- [ ] `UpdateMeasure_WithValidFormula_UpdatesSuccessfully` (WRITE - keep fresh)

**File: `DataModelCommandsTests.Relationships.cs`**
- [ ] `CreateRelationship_WithValidParameters_CreatesSuccessfully` (WRITE - keep fresh)
- [ ] `DeleteRelationship_WithValidRelationship_ReturnsSuccessResult` (WRITE - keep fresh)

### Step 2: Update READ-ONLY Tests

**Pattern: Add `requiresWritableDataModel: false` parameter**

#### File: `DataModelCommandsTests.Tables.cs`

```csharp
[Fact]
public async Task ListTables_WithRealisticDataModel_ReturnsTablesWithData()
{
    // Arrange - Use template (fast - just file copy!)
    var testFile = await CreateTestFileAsync(
        "ListTables_WithRealisticDataModel_ReturnsTablesWithData.xlsx", 
        requiresWritableDataModel: false);  // 👈 ADD THIS

    // Rest of test unchanged...
}

[Fact]
public async Task ListTables_WithValidFile_ReturnsSuccessResult()
{
    // Arrange - Use template (fast - just file copy!)
    var testFile = await CreateTestFileAsync(
        "ListTables_WithValidFile_ReturnsSuccessResult.xlsx",
        requiresWritableDataModel: false);  // 👈 ADD THIS

    // Rest of test unchanged...
}
```

#### File: `DataModelCommandsTests.Discovery.cs`

```csharp
[Fact]
public async Task ListTableColumns_WithValidTable_ReturnsColumns()
{
    var testFile = await CreateTestFileAsync(
        "ListTableColumns_WithValidTable_ReturnsColumns.xlsx",
        requiresWritableDataModel: false);  // 👈 ADD THIS
    // ...
}

[Fact]
public async Task ViewTable_WithValidTable_ReturnsCompleteInfo()
{
    var testFile = await CreateTestFileAsync(
        "ViewTable_WithValidTable_ReturnsCompleteInfo.xlsx",
        requiresWritableDataModel: false);  // 👈 ADD THIS
    // ...
}

[Fact]
public async Task ViewTable_WithTableHavingMeasures_CountsMeasuresCorrectly()
{
    var testFile = await CreateTestFileAsync(
        "ViewTable_WithTableHavingMeasures_CountsMeasuresCorrectly.xlsx",
        requiresWritableDataModel: false);  // 👈 ADD THIS
    // ...
}

[Fact]
public async Task GetModelInfo_WithRealisticDataModel_ReturnsAccurateStatistics()
{
    var testFile = await CreateTestFileAsync(
        "GetModelInfo_WithRealisticDataModel_ReturnsAccurateStatistics.xlsx",
        requiresWritableDataModel: false);  // 👈 ADD THIS
    // ...
}
```

#### File: `DataModelCommandsTests.Measures.cs`

```csharp
[Fact]
public async Task ListMeasures_WithRealisticDataModel_ReturnsMeasuresWithFormulas()
{
    var testFile = await CreateTestFileAsync(
        "ListMeasures_WithRealisticDataModel_ReturnsMeasuresWithFormulas.xlsx",
        requiresWritableDataModel: false);  // 👈 ADD THIS
    // ...
}

[Fact]
public async Task ViewMeasure_WithRealisticDataModel_ReturnsValidDAXFormula()
{
    var testFile = await CreateTestFileAsync(
        "ViewMeasure_WithRealisticDataModel_ReturnsValidDAXFormula.xlsx",
        requiresWritableDataModel: false);  // 👈 ADD THIS
    // ...
}

// LEAVE THESE UNCHANGED (WRITE operations):
// - CreateMeasure_WithValidParameters_CreatesSuccessfully
// - CreateMeasure_WithFormatType_CreatesWithFormat
// - UpdateMeasure_WithValidFormula_UpdatesSuccessfully
```

#### File: `DataModelCommandsTests.Relationships.cs`

```csharp
[Fact]
public async Task ListRelationships_WithValidFile_ReturnsSuccessResult()
{
    var testFile = await CreateTestFileAsync(
        "ListRelationships_WithValidFile_ReturnsSuccessResult.xlsx",
        requiresWritableDataModel: false);  // 👈 ADD THIS
    // ...
}

[Fact]
public async Task ListRelationships_WithRealisticDataModel_ReturnsRelationshipsWithTables()
{
    var testFile = await CreateTestFileAsync(
        "ListRelationships_WithRealisticDataModel_ReturnsRelationshipsWithTables.xlsx",
        requiresWritableDataModel: false);  // 👈 ADD THIS
    // ...
}

[Fact]
public async Task ViewRelationship_WithValidRelationship_ReturnsCompleteInfo()
{
    var testFile = await CreateTestFileAsync(
        "ViewRelationship_WithValidRelationship_ReturnsCompleteInfo.xlsx",
        requiresWritableDataModel: false);  // 👈 ADD THIS
    // ...
}

[Fact]
public async Task GetRelationshipProperties_WithActiveRelationship_ReturnsActiveState()
{
    var testFile = await CreateTestFileAsync(
        "GetRelationshipProperties_WithActiveRelationship_ReturnsActiveState.xlsx",
        requiresWritableDataModel: false);  // 👈 ADD THIS
    // ...
}

[Fact]
public async Task GetRelationshipProperties_WithInactiveRelationship_ReturnsInactiveState()
{
    var testFile = await CreateTestFileAsync(
        "GetRelationshipProperties_WithInactiveRelationship_ReturnsInactiveState.xlsx",
        requiresWritableDataModel: false);  // 👈 ADD THIS
    // ...
}

// LEAVE THESE UNCHANGED (WRITE operations):
// - CreateRelationship_WithValidParameters_CreatesSuccessfully
// - DeleteRelationship_WithValidRelationship_ReturnsSuccessResult
```

### Step 3: Verify Implementation

**Run tests individually first:**

```powershell
cd D:\source\mcp-server-excel

# Test a single READ-ONLY test
dotnet test --filter "FullyQualifiedName~ListTables_WithRealisticDataModel_ReturnsTablesWithData" -v detailed

# Should see: Template copied (fast ~1-2s) instead of fresh build (slow ~7-10s)
```

**Run all DataModel tests:**

```powershell
# Measure before optimization
$sw = [System.Diagnostics.Stopwatch]::StartNew()
dotnet test --filter "Feature=DataModel" --no-build
$sw.Stop()
Write-Output "Time: $($sw.Elapsed.TotalSeconds) seconds"

# Expected BEFORE: ~120 seconds
# Expected AFTER: ~55 seconds (13 tests × 2s + 3 tests × 10s)
```

### Step 4: Validate Template Creation

**Ensure template exists and is valid:**

```powershell
$templatePath = "tests\ExcelMcp.Core.Tests\bin\Debug\net8.0\TestAssets\DataModelTemplate.xlsx"

if (Test-Path $templatePath) {
    $info = Get-Item $templatePath
    Write-Output "✅ Template exists: $($info.FullName)"
    Write-Output "   Size: $($info.Length) bytes"
    Write-Output "   Created: $($info.CreationTime)"
    Write-Output "   Modified: $($info.LastWriteTime)"
} else {
    Write-Output "❌ Template missing - will be created on first test run"
}
```

**If template doesn't exist, it will be created automatically on first READ-ONLY test run (60-120s one-time cost).**

### Step 5: Performance Verification

**Expected Results:**

| Metric | Before | After | Improvement |
|--------|--------|-------|-------------|
| Total time | 120s | 55s | 54% faster |
| READ-ONLY test (avg) | 7.5s | 2s | 73% faster |
| WRITE test (avg) | 7.5s | 10s | 33% slower* |
| Template creation | N/A | 60-120s | One-time cost |

*WRITE tests unchanged, but appear slower relatively because READ tests got faster.

**Success Criteria:**
- ✅ All 16 tests pass
- ✅ Total execution time < 60 seconds
- ✅ Template file created in TestAssets folder
- ✅ No test failures or regressions

### Step 6: Commit Changes

```bash
# Verify changes
git status
git diff tests/ExcelMcp.Core.Tests/Integration/Commands/DataModel/

# Commit
git add tests/ExcelMcp.Core.Tests/Integration/Commands/DataModel/*.cs
git commit -m "test(datamodel): optimize READ-ONLY tests with template files

- Add requiresWritableDataModel: false to 13 READ-ONLY tests
- Tests now copy template instead of building fresh Data Model
- Performance: 120s → 55s (54% improvement)
- Template created once on first test run, reused thereafter

Refs: docs/DATAMODEL-TEST-OPTIMIZATION.md"
```

## Troubleshooting

### Issue: Template Not Created

**Symptom:** Tests still take 7-10s each

**Solution:**
```powershell
# Delete any existing template and force recreation
Remove-Item "tests\ExcelMcp.Core.Tests\bin\Debug\net8.0\TestAssets\DataModelTemplate.xlsx" -ErrorAction SilentlyContinue

# Run a single READ-ONLY test to create template
dotnet test --filter "FullyQualifiedName~ListTables_WithRealisticDataModel_ReturnsTablesWithData"

# Verify template created
Test-Path "tests\ExcelMcp.Core.Tests\bin\Debug\net8.0\TestAssets\DataModelTemplate.xlsx"
```

### Issue: Tests Fail After Template

**Symptom:** Tests pass before but fail after adding `requiresWritableDataModel: false`

**Diagnosis:**
- Template might be corrupted
- Template might not have expected Data Model structure

**Solution:**
```powershell
# Recreate template
Remove-Item "tests\ExcelMcp.Core.Tests\bin\Debug\net8.0\TestAssets\DataModelTemplate.xlsx"

# Run test to recreate
dotnet test --filter "FullyQualifiedName~ListTables_WithRealisticDataModel_ReturnsTablesWithData"

# If still fails, inspect template manually in Excel
start "tests\ExcelMcp.Core.Tests\bin\Debug\net8.0\TestAssets\DataModelTemplate.xlsx"
```

### Issue: File Lock Errors

**Symptom:** `IOException: The process cannot access the file because it is being used by another process`

**Solution:**
```powershell
# Close all Excel instances
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force

# Clear temp directory
Remove-Item "C:\Users\[USER]\AppData\Local\Temp\DataModelCommandsTests_*" -Recurse -Force -ErrorAction SilentlyContinue

# Retry tests
dotnet test --filter "Feature=DataModel"
```

## Next Steps (Week 2)

After Week 1 optimization is complete and verified:

1. **Create static test asset files** (Week 2)
   - Pre-built Data Model files in TestAssets
   - Eliminates template creation overhead
   - Additional 13s savings

2. **Optimize write test fixture** (Week 3)
   - Minimal Data Model in fixture
   - Reduce fixture setup from 60-120s to 10-15s
   - Additional 15s savings

3. **Document maintenance process** (Week 4)
   - Asset regeneration script
   - CI/CD validation
   - Developer guidelines

---

## Summary

**Changes:** Add one parameter (`requiresWritableDataModel: false`) to 13 test methods

**Time Required:** 30-60 minutes

**Performance Gain:** 65 seconds (54% improvement)

**Risk:** Low - template strategy already implemented, just needs to be enabled
