# DataModel Test Optimization Analysis

## Executive Summary

**Current Performance:** 16 tests in ~120 seconds = **7.5 seconds/test** (SLOW!)

**Root Cause:** Every test creates a fresh Data Model from scratch (3 worksheets + 3 tables + 2 relationships + 3 measures)

**Optimization Potential:** Can reduce to ~30-40 seconds total (67-75% improvement)

---

## Current Architecture

### Test Structure (16 tests total)

```
DataModelCommandsTests (partial class split across 5 files)
├── DataModelCommandsTests.cs (1 test - base class)
├── DataModelCommandsTests.Discovery.cs (4 tests)
├── DataModelCommandsTests.Measures.cs (6 tests)
├── DataModelCommandsTests.Relationships.cs (4 tests)
└── DataModelCommandsTests.Tables.cs (2 tests)

DataModelWriteTests (separate class with shared fixture)
└── Uses DataModelWriteTestsFixture (3 tests - currently has issues)
```

### Current Pattern (SLOW - 7.5s per test)

**Each test calls:**
```csharp
var testFile = await CreateTestFileAsync("TestName.xlsx");
```

**Which creates:**
1. Empty workbook (2-3s)
2. Sales worksheet with 10 rows + format as Table (1-2s)
3. Customers worksheet with 5 rows + format as Table (1-2s)
4. Products worksheet with 5 rows + format as Table (1-2s)
5. Add 3 tables to Data Model (0.5s each = 1.5s)
6. Create 2 relationships (0.5s each = 1s)
7. Create 3 measures (0.5s each = 1.5s)

**Total per test: 10-15 seconds of setup** (most tests only need 1-2s of actual testing!)

### Template Strategy (partially implemented)

Code supports template strategy but **NO tests actually use it!**

```csharp
// NEVER CALLED - requiresWritableDataModel flag not used by any test
var testFile = await CreateTestFileAsync("Test.xlsx", requiresWritableDataModel: false);
```

**Why?** All current tests just call:
```csharp
await CreateTestFileAsync("TestName.xlsx")  // Defaults to fresh build
```

---

## Optimization Strategy

### Phase 1: Template-Based Tests (QUICK WIN)

**Goal:** Reduce READ-ONLY tests from 7.5s to 2-3s each

**Approach:**
1. ✅ Template file already exists in `bin/Debug/net8.0/TestAssets/DataModelTemplate.xlsx`
2. ✅ Code supports template copying
3. ❌ No tests actually use it!

**Solution:** Mark 13 READ-ONLY tests to use template:

```csharp
// BEFORE (7.5s)
var testFile = await CreateTestFileAsync("ListTables.xlsx");

// AFTER (1-2s - just file copy!)
var testFile = await CreateTestFileAsync("ListTables.xlsx", requiresWritableDataModel: false);
```

**READ-ONLY Tests (13 total):**
- `ListTables_WithRealisticDataModel_ReturnsTablesWithData`
- `ListTables_WithValidFile_ReturnsSuccessResult`
- `ListTableColumns_WithValidTable_ReturnsColumns`
- `ViewTable_WithValidTable_ReturnsCompleteInfo`
- `ViewTable_WithTableHavingMeasures_CountsMeasuresCorrectly`
- `GetModelInfo_WithRealisticDataModel_ReturnsAccurateStatistics`
- `ListMeasures_WithRealisticDataModel_ReturnsMeasuresWithFormulas`
- `ViewMeasure_WithRealisticDataModel_ReturnsValidDAXFormula`
- `ListRelationships_WithValidFile_ReturnsSuccessResult`
- `ListRelationships_WithRealisticDataModel_ReturnsRelationshipsWithTables`
- `ViewRelationship_WithValidRelationship_ReturnsCompleteInfo`
- `GetRelationshipProperties_WithActiveRelationship_ReturnsActiveState`
- `GetRelationshipProperties_WithInactiveRelationship_ReturnsInactiveState`

**WRITE Tests (3 total - keep fresh build):**
- `CreateMeasure_WithValidParameters_CreatesSuccessfully`
- `CreateMeasure_WithFormatType_CreatesWithFormat`
- `UpdateMeasure_WithValidFormula_UpdatesSuccessfully`
- `CreateRelationship_WithValidParameters_CreatesSuccessfully`
- `DeleteRelationship_WithValidRelationship_ReturnsSuccessResult`

**Expected Savings:**
- 13 tests × 5s savings = **65 seconds saved**
- New total: ~55 seconds (13 × 2s + 3 × 10s)

---

### Phase 2: Static Test Asset Files (BETTER)

**Goal:** Eliminate template creation overhead entirely

**Current Issue:** Template created on first test run (60-120s one-time cost)

**Solution:** Pre-built static test files in `tests/ExcelMcp.Core.Tests/TestAssets/`

#### Recommended Static Files

1. **`DataModel-Basic.xlsx`** (minimal Data Model)
   - Sales table (5 rows)
   - 1 measure: `Total Sales = SUM(Sales[Amount])`
   - Use for: Basic list/view operations

2. **`DataModel-Relationships.xlsx`** (multi-table)
   - Sales table (5 rows)
   - Customers table (3 rows)
   - Products table (3 rows)
   - 2 relationships (Sales→Customers, Sales→Products)
   - Use for: Relationship tests

3. **`DataModel-Measures.xlsx`** (measure-heavy)
   - Sales table (5 rows)
   - 5 measures with various DAX formulas
   - Use for: Measure discovery tests

4. **`DataModel-Complete.xlsx`** (comprehensive)
   - All tables, relationships, measures
   - Use for: GetModelInfo, complex scenarios

#### Benefits
- ✅ Zero template creation time
- ✅ Version controlled (git tracks changes)
- ✅ Reproducible across environments
- ✅ Smaller files (5 rows vs 10 rows = faster copy)
- ✅ Can test against specific Data Model states

#### Implementation Process

**Step 1: Create Static Files**
```powershell
# Run helper script to generate static test files
dotnet run -p tests/ExcelMcp.Core.Tests/TestAssets/CreateDataModelAsset.cs
```

**Step 2: Update Test Class**
```csharp
private async Task<string> CreateTestFileAsync(string fileName, 
    DataModelAssetType assetType = DataModelAssetType.Complete)
{
    var templateName = assetType switch
    {
        DataModelAssetType.Basic => "DataModel-Basic.xlsx",
        DataModelAssetType.Relationships => "DataModel-Relationships.xlsx",
        DataModelAssetType.Measures => "DataModel-Measures.xlsx",
        _ => "DataModel-Complete.xlsx"
    };
    
    var templatePath = Path.Combine(
        Path.GetDirectoryName(typeof(DataModelCommandsTests).Assembly.Location)!,
        "TestAssets", templateName);
    
    var testFilePath = Path.Combine(_tempDir, fileName);
    File.Copy(templatePath, testFilePath, overwrite: true);
    
    // Ensure writable
    new FileInfo(testFilePath).IsReadOnly = false;
    
    return testFilePath;
}
```

**Step 3: Update .csproj**
```xml
<ItemGroup>
  <None Include="TestAssets\DataModel-*.xlsx">
    <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
  </None>
</ItemGroup>
```

**Expected Savings:**
- 13 tests × 6s savings = **78 seconds saved**
- New total: ~42 seconds (13 × 1s + 3 × 10s)

---

### Phase 3: Shared Fixture Optimization (ADVANCED)

**Current DataModelWriteTestsFixture Issues:**
1. Creates full Data Model (60-120s)
2. Only used by 3 tests in DataModelWriteTests class
3. Creates measures in fixture, then tests try to create more → COM errors

**Better Approach: Minimal Fixture**

```csharp
public class DataModelWriteTestsFixture : IAsyncLifetime
{
    public async Task InitializeAsync()
    {
        // Create MINIMAL Data Model (10-15s)
        // - Sales table (3 rows) - for measure tests
        // - Customers table (2 rows) - for relationship tests
        // - NO relationships
        // - NO measures
        // Let tests create their own objects
        
        await CreateMinimalDataModelAsync();
    }
}
```

**Benefits:**
- Fixture setup: 60-120s → 10-15s (80-90% faster)
- Tests can create measures/relationships without conflicts
- Total write tests: 3 × 5s = 15s (vs current 3 × 10s = 30s)

---

### Phase 4: Parallel Test Execution (FUTURE)

**Current:** xUnit runs tests sequentially in same class

**Optimization:** Split into multiple test classes

```
DataModelDiscoveryTests (IClassFixture<TempDirectoryFixture>)
DataModelMeasureTests (IClassFixture<TempDirectoryFixture>)
DataModelRelationshipTests (IClassFixture<TempDirectoryFixture>)
DataModelTableTests (IClassFixture<TempDirectoryFixture>)
```

**Benefits:**
- xUnit runs different classes in parallel
- Each gets own fixture/temp directory
- Could reduce total time by 50-75% if 4-way parallelism

**Risks:**
- Excel COM is STA (single-threaded apartment)
- Multiple Excel instances might cause issues
- Need testing to verify stability

---

## Process to Keep Static Files Current

### Strategy 1: Asset Builder Script (RECOMMENDED)

**Create:** `tests/ExcelMcp.Core.Tests/TestAssets/CreateDataModelAssets.ps1`

```powershell
# DataModel Test Asset Builder
# Generates static test Excel files with Data Models

$ErrorActionPreference = "Stop"
$projectPath = "D:\source\mcp-server-excel"
$assetsDir = "$projectPath\tests\ExcelMcp.Core.Tests\TestAssets"

Write-Host "Creating DataModel test assets..." -ForegroundColor Cyan

# Use C# script to create assets via production commands
dotnet fsi "$assetsDir\CreateDataModelAsset.cs" `
    --output "$assetsDir\DataModel-Basic.xlsx" `
    --type Basic

dotnet fsi "$assetsDir\CreateDataModelAsset.cs" `
    --output "$assetsDir\DataModel-Relationships.xlsx" `
    --type Relationships

dotnet fsi "$assetsDir\CreateDataModelAsset.cs" `
    --output "$assetsDir\DataModel-Measures.xlsx" `
    --type Measures

dotnet fsi "$assetsDir\CreateDataModelAsset.cs" `
    --output "$assetsDir\DataModel-Complete.xlsx" `
    --type Complete

Write-Host "✅ Assets created successfully" -ForegroundColor Green
Write-Host "Files: $assetsDir\DataModel-*.xlsx" -ForegroundColor Gray
```

**Usage:**
```powershell
# Rebuild all test assets
cd tests\ExcelMcp.Core.Tests\TestAssets
.\CreateDataModelAssets.ps1

# Commit updated files
git add DataModel-*.xlsx
git commit -m "test: update DataModel test assets"
```

### Strategy 2: CI/CD Validation (AUTOMATED)

**Add to `.github/workflows/test.yml`:**

```yaml
- name: Validate DataModel Test Assets
  run: |
    # Check assets exist
    if (!(Test-Path "tests\ExcelMcp.Core.Tests\TestAssets\DataModel-*.xlsx")) {
      Write-Error "DataModel test assets missing!"
      exit 1
    }
    
    # Check assets are not too old (older than 90 days)
    Get-ChildItem "tests\ExcelMcp.Core.Tests\TestAssets\DataModel-*.xlsx" | ForEach-Object {
      $age = (Get-Date) - $_.LastWriteTime
      if ($age.Days -gt 90) {
        Write-Warning "$($_.Name) is $($age.Days) days old - consider regenerating"
      }
    }
```

### Strategy 3: Developer Guidelines

**Add to `docs/CONTRIBUTING.md`:**

```markdown
### DataModel Test Asset Maintenance

When modifying Data Model commands, verify test assets are current:

1. **Check asset age:**
   ```powershell
   ls tests\ExcelMcp.Core.Tests\TestAssets\DataModel-*.xlsx | 
       select Name, LastWriteTime
   ```

2. **Regenerate if needed:**
   ```powershell
   cd tests\ExcelMcp.Core.Tests\TestAssets
   .\CreateDataModelAssets.ps1
   ```

3. **Commit changes:**
   ```bash
   git add TestAssets/DataModel-*.xlsx
   git commit -m "test: regenerate DataModel assets after command changes"
   ```

**When to regenerate:**
- After changing Data Model commands
- After modifying table/measure/relationship creation logic
- Every 90 days (automated reminder in CI/CD)
- Before major releases
```

### Strategy 4: Asset Version Metadata (ADVANCED)

**Embed metadata in test assets:**

```csharp
// In CreateDataModelAsset.cs
private static async Task EmbedMetadataAsync(string filePath)
{
    await using var batch = await ExcelSession.BeginBatchAsync(filePath);
    await batch.ExecuteAsync<int>((ctx, ct) =>
    {
        // Add hidden metadata worksheet
        dynamic sheet = ctx.Book.Worksheets.Add();
        sheet.Name = "_Metadata";
        sheet.Visible = -4159; // xlSheetVeryHidden
        
        sheet.Range["A1"].Value2 = "AssetVersion";
        sheet.Range["B1"].Value2 = "1.0.0"; // Semantic version
        
        sheet.Range["A2"].Value2 = "CreatedDate";
        sheet.Range["B2"].Value2 = DateTime.UtcNow.ToString("o");
        
        sheet.Range["A3"].Value2 = "CreatedBy";
        sheet.Range["B3"].Value2 = "CreateDataModelAsset.cs";
        
        sheet.Range["A4"].Value2 = "Description";
        sheet.Range["B4"].Value2 = "Test asset for DataModel integration tests";
        
        return ValueTask.FromResult(0);
    });
    await batch.SaveAsync();
}
```

**Validation in tests:**

```csharp
[Fact]
public async Task DataModelAssets_HaveCurrentMetadata()
{
    var assetsDir = Path.Combine(
        Path.GetDirectoryName(typeof(DataModelCommandsTests).Assembly.Location)!,
        "TestAssets");
    
    var assets = Directory.GetFiles(assetsDir, "DataModel-*.xlsx");
    
    foreach (var asset in assets)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(asset);
        var version = await ReadAssetVersionAsync(batch);
        
        Assert.Equal("1.0.0", version, 
            $"{Path.GetFileName(asset)} has outdated version {version}");
    }
}
```

---

## Recommended Implementation Order

### Week 1: Quick Wins (65s savings)
1. ✅ Mark 13 READ-ONLY tests with `requiresWritableDataModel: false`
2. ✅ Verify template strategy works
3. ✅ Measure new performance

**Expected result:** 120s → 55s (54% improvement)

### Week 2: Static Assets (additional 13s savings)
1. ✅ Create `CreateDataModelAssets.ps1` script
2. ✅ Generate 4 static test files
3. ✅ Update `.csproj` to copy assets
4. ✅ Update tests to use specific assets
5. ✅ Add asset validation to CI/CD

**Expected result:** 55s → 42s (65% improvement total)

### Week 3: Fixture Optimization (additional 15s savings)
1. ✅ Simplify `DataModelWriteTestsFixture` to minimal Data Model
2. ✅ Update write tests to create their own measures/relationships
3. ✅ Measure performance

**Expected result:** 42s → 27s (78% improvement total)

### Week 4: Documentation & Maintenance
1. ✅ Document asset creation process
2. ✅ Add developer guidelines
3. ✅ Create asset regeneration schedule
4. ✅ Add metadata versioning (optional)

---

## Performance Comparison

| Phase | Tests | Time | Savings | Notes |
|-------|-------|------|---------|-------|
| **Current** | 16 | 120s | - | Fresh build every test |
| **Phase 1** (Template) | 13+3 | 55s | 54% | Template copy for reads |
| **Phase 2** (Static) | 13+3 | 42s | 65% | Pre-built static files |
| **Phase 3** (Fixture) | 13+3 | 27s | 78% | Minimal fixture setup |
| **Phase 4** (Parallel) | 16 | 10-15s | 87-92% | Multi-class parallelism (risky) |

---

## Risks & Mitigations

### Risk 1: Template/Asset Staleness
**Mitigation:**
- CI/CD age validation (warn if >90 days)
- Developer guidelines
- Asset regeneration script

### Risk 2: Test Dependencies
**Mitigation:**
- Each test uses unique measure/relationship names
- Template is read-only (copied per test)
- Verify tests pass individually and together

### Risk 3: Platform Differences
**Mitigation:**
- Static assets created on Windows with Excel
- Include metadata about creation environment
- CI/CD validates assets exist and are valid

### Risk 4: Parallel Execution Issues
**Mitigation:**
- Phase 4 is optional
- Test thoroughly before enabling
- Fallback to sequential if issues

---

## Conclusion

**Immediate Action (Week 1):**
Add `requiresWritableDataModel: false` to 13 READ-ONLY tests → **65s savings (54% improvement)**

**Best Long-Term Solution:**
Static test assets + minimal fixture → **93s savings (78% improvement)**

**Maintenance:**
Asset builder script + CI/CD validation + 90-day regeneration schedule

**Total Estimated Effort:**
- Week 1: 2-3 hours (mark tests, verify)
- Week 2: 4-6 hours (create assets, update tests)
- Week 3: 2-3 hours (optimize fixture)
- Week 4: 2-3 hours (documentation)
- **Total: 10-15 hours for 78% performance improvement**
