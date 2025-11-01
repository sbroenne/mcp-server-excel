# Data Model Test Optimization Strategy

## Current Situation (Slow)

### DataModelWriteTests
- ✅ Uses shared fixture (`DataModelWriteTestsFixture`)
- Creates ONE file for ALL write tests (60-120s total)
- **Good pattern** - multiple tests share one setup

### DataModelCommandsTests (READ tests)
- ❌ Each test creates its own Data Model from scratch
- ~10 tests × 10-15s each = **100-150 seconds total**
- **Bad pattern** - recreating same structure repeatedly

## Optimization Solution

### Option 1: Static Pre-Built Asset File (RECOMMENDED)

**Approach:**
1. Create `tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx` (pre-built, committed to repo)
2. Use `DataModelAssetBuilder` to generate it once
3. READ tests copy this template at test start
4. WRITE tests continue using `DataModelWriteTestsFixture` (fresh file per run)

**Benefits:**
- ✅ **90% faster READ tests** (copy file ~0.5s vs create ~10s per test)
- ✅ Template checked into repo = consistent across environments
- ✅ CI doesn't need to build Data Model each time
- ✅ Template can be regenerated when schema changes

**Strategy to Keep Current:**
1. Add build target: `dotnet run --project tests/ExcelMcp.Core.Tests/TestAssets/BuildAssets.csproj`
2. Manual regeneration: When Data Model schema changes, run builder and commit new template
3. Version check: Add hash/timestamp in file properties to detect outdated template
4. CI verification: Test that fails if template is outdated (compares schema)

**Implementation:**
```csharp
// New fixture for READ tests
public class DataModelReadTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;
    public string TestFilePath { get; private set; } = null!;
    
    public async Task InitializeAsync()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"DataModelReadTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        TestFilePath = Path.Join(_tempDir, "DataModel.xlsx");
        
        // Copy pre-built template (fast!)
        var templatePath = "TestAssets/DataModelTemplate.xlsx";
        File.Copy(templatePath, TestFilePath, overwrite: true);
        
        // ~0.5s vs 60-120s!
    }
    
    public Task DisposeAsync()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
        return Task.CompletedTask;
    }
}
```

**Regeneration Script:**
```powershell
# scripts/regenerate-test-assets.ps1
dotnet run --project tests/ExcelMcp.Core.Tests --no-build -- generate-datamodel-asset
git add tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx
git commit -m "test: Regenerate Data Model test asset"
```

### Option 2: Shared Fixture for READ Tests

**Approach:**
- Use `IClassFixture<DataModelReadTestsFixture>` for all READ tests
- Create Data Model ONCE, all READ tests share it
- Same pattern as WRITE tests

**Benefits:**
- ✅ **80% faster** (60-120s once vs per test)
- ✅ No need to commit binary file

**Drawbacks:**
- ❌ READ tests become interdependent (fixture lifecycle)
- ❌ Still slow in CI (must rebuild each run)
- ❌ Tests can't run in parallel (share same file)

### Option 3: Lazy Static Fixture

**Approach:**
- Singleton fixture that creates Data Model ONCE per test session
- All tests use the same file (read-only)
- xUnit Collection Fixture pattern

**Benefits:**
- ✅ Build once, use by ALL tests
- ✅ Tests can run in parallel (read-only access)

**Drawbacks:**
- ❌ Complex setup (collection fixtures)
- ❌ Tests can interfere if they modify (need isolation)

## Recommendation

**Use Option 1: Pre-Built Static Asset**

### Why:
1. **Fastest**: Copy file (0.5s) vs build (60-120s)
2. **CI-friendly**: No Excel operations needed for setup
3. **Consistent**: Same data across all environments
4. **Maintainable**: Explicit regeneration when schema changes
5. **Parallel-safe**: Each test gets its own copy

### Implementation Plan:

#### Step 1: Generate Template
```bash
# Run builder to create template
dotnet run --project tests/ExcelMcp.Core.Tests -- --build-asset datamodel

# Verify template
ls -l tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx
```

#### Step 2: Create Read Fixture
```csharp
// tests/ExcelMcp.Core.Tests/Helpers/DataModelReadTestsFixture.cs
public class DataModelReadTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;
    public string TestFilePath { get; private set; } = null!;
    
    public Task InitializeAsync()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"DataModelReadTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        TestFilePath = Path.Join(_tempDir, "DataModel.xlsx");
        
        // Copy template (FAST!)
        var solutionRoot = Path.GetFullPath(Path.Join(AppContext.BaseDirectory, "../../../../.."));
        var templatePath = Path.Join(solutionRoot, "tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx");
        
        if (!File.Exists(templatePath))
        {
            throw new FileNotFoundException($"Data Model template not found. Run: dotnet test --filter BuildAssets. Path: {templatePath}");
        }
        
        File.Copy(templatePath, TestFilePath, overwrite: true);
        return Task.CompletedTask;
    }
    
    public Task DisposeAsync()
    {
        if (Directory.Exists(_tempDir)) Directory.Delete(_tempDir, recursive: true);
        return Task.CompletedTask;
    }
}
```

#### Step 3: Update READ Tests
```csharp
// Change from TempDirectoryFixture to DataModelReadTestsFixture
public partial class DataModelCommandsTests : IClassFixture<DataModelReadTestsFixture>
{
    private readonly DataModelCommands _commands;
    private readonly string _testFilePath; // from fixture
    
    public DataModelCommandsTests(DataModelReadTestsFixture fixture)
    {
        _commands = new DataModelCommands();
        _testFilePath = fixture.TestFilePath;
    }
    
    [Fact]
    public async Task ListTables_WithValidFile_ReturnsSuccessResult()
    {
        // Use pre-built file from fixture (already has Data Model!)
        await using var batch = await ExcelSession.BeginBatchAsync(_testFilePath);
        var result = await _commands.ListTablesAsync(batch);
        
        Assert.True(result.Success);
        Assert.NotEmpty(result.Tables);
        // Fast! No Data Model creation needed!
    }
}
```

#### Step 4: Template Versioning
```csharp
// Add to DataModelAssetBuilder
private const string ASSET_VERSION = "1.0.0";

public static async Task<string> CreateDataModelAssetAsync(string targetPath)
{
    // ... create tables, relationships, measures ...
    
    // Add version metadata
    await batch.Execute<int>((ctx, ct) =>
    {
        dynamic props = ctx.Book.BuiltinDocumentProperties;
        props.Item("Comments").Value = $"DataModelTemplate v{ASSET_VERSION} - Generated {DateTime.UtcNow:O}";
        return 0;
    });
    
    await batch.SaveAsync();
}
```

### Maintenance Strategy

**When to Regenerate:**
1. Data Model schema changes (new table, relationship, measure)
2. Test requirements change (need different data)
3. Quarterly maintenance (update dates, refresh data)

**How to Regenerate:**
```bash
# 1. Update CreateDataModelAsset.cs with new schema
# 2. Run generator
dotnet run --project tests/ExcelMcp.Core.Tests -- --build-asset datamodel

# 3. Commit new template
git add tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx
git commit -m "test: Update Data Model template (v1.1.0 - added XYZ table)"
```

**CI Verification:**
```csharp
[Fact]
public async Task DataModelTemplate_IsCurrentVersion()
{
    var templatePath = "TestAssets/DataModelTemplate.xlsx";
    await using var batch = await ExcelSession.BeginBatchAsync(templatePath);
    
    var version = await batch.Execute<string>((ctx, ct) =>
    {
        dynamic props = ctx.Book.BuiltinDocumentProperties;
        return props.Item("Comments").Value?.ToString() ?? "";
    });
    
    Assert.Contains(DataModelAssetBuilder.ASSET_VERSION, version);
}
```

## Expected Performance Improvement

| Scenario | Before | After | Improvement |
|----------|--------|-------|-------------|
| Single READ test | ~10s | ~0.5s | **95% faster** |
| 10 READ tests | ~100s | ~5s | **95% faster** |
| WRITE tests | 60-120s | 60-120s | No change (already optimal) |
| **Total test suite** | ~160-220s | ~65-125s | **60% faster** |

## Risk Mitigation

**Risk**: Template becomes outdated
- **Mitigation**: Version check test fails if mismatch
- **Mitigation**: Document regeneration process in CONTRIBUTING.md
- **Mitigation**: Pre-commit hook warns if CreateDataModelAsset.cs changed without template

**Risk**: Template corrupted
- **Mitigation**: CI test verifies template structure
- **Mitigation**: Keep generator in repo to recreate

**Risk**: Environment-specific issues
- **Mitigation**: Template is Excel COM API output (portable)
- **Mitigation**: CI rebuilds if template missing
