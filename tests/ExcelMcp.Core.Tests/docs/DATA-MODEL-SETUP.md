# Data Model Test Setup

> **Current implementation for fast Data Model integration tests**

## Overview

Data Model tests use a **pre-built template file** for fast setup:
- **Template creation**: ~60-120 seconds (one-time)
- **Template copy**: ~0.5 seconds per test
- **Performance gain**: 95% faster than building from scratch

## Current Architecture

### Test Fixtures

**DataModelReadTestsFixture** - For READ-only tests (list, view, get)
- Copies `TestAssets/DataModelTemplate.xlsx` to temp directory
- Each test class gets its own copy
- Fast setup (~0.5s vs 60-120s)

**DataModelWriteTestsFixture** - For WRITE tests (create, update, delete)
- Creates fresh Data Model file
- Slower setup (60-120s) but necessary for isolation
- One file shared by all write tests in the class

### Template File

**Location**: `tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx`

**Contents**:
- 3 Excel Tables: Sales, Customers, Products
- 2 Relationships: Sales→Customers, Sales→Products
- 3 DAX Measures: Total Sales, Average Sale, Total Customers

**Status**: ✅ Template file is stored in git and ready to use

## Using the Test Fixtures

### For READ Tests (List, View, Get)

```csharp
[Trait("Category", "Integration")]
[Trait("Feature", "DataModel")]
public partial class DataModelCommandsTests : IClassFixture<DataModelReadTestsFixture>
{
    private readonly DataModelCommands _commands;
    private readonly string _testFilePath;

    public DataModelCommandsTests(DataModelReadTestsFixture fixture)
    {
        _commands = new DataModelCommands();
        _testFilePath = fixture.TestFilePath;  // Pre-built template copy
    }

    [Fact]
    public async Task ListTables_ReturnsExpectedTables()
    {
        // Fast! Uses pre-built template (0.5s setup)
        await using var batch = await ExcelSession.BeginBatchAsync(_testFilePath);
        var result = await _commands.ListTablesAsync(batch);
        
        Assert.True(result.Success);
        Assert.Equal(3, result.Tables.Count);  // Sales, Customers, Products
    }
}
```

### For WRITE Tests (Create, Update, Delete)

```csharp
[Trait("Category", "Integration")]
[Trait("Feature", "DataModel")]
public partial class DataModelWriteTests : IClassFixture<DataModelWriteTestsFixture>
{
    private readonly DataModelCommands _commands;
    private readonly string _testFilePath;

    public DataModelWriteTests(DataModelWriteTestsFixture fixture)
    {
        _commands = new DataModelCommands();
        _testFilePath = fixture.TestFilePath;  // Fresh file
    }

    [Fact]
    public async Task CreateMeasure_ValidDax_CreatesSuccessfully()
    {
        // Slower setup (60-120s) but fresh file ensures isolation
        await using var batch = await ExcelSession.BeginBatchAsync(_testFilePath);
        var result = await _commands.CreateMeasureAsync(
            batch, "Sales", "TestMeasure", "SUM(Sales[Amount])", "Currency");
        
        Assert.True(result.Success);
    }
}
```

## Generating the Template

### 📝 Template is Stored in Git

The template file (`DataModelTemplate.xlsx`) is checked into git and rarely needs regeneration.

**Only regenerate if:**
- You need to change the table structure
- You need to add/modify relationships
- You need to add/modify measures

### How to Regenerate (Rarely Needed)

```bash
# Run from tests/ExcelMcp.Core.Tests directory

# 1. Build the test project (needed for the builder)
dotnet build -c Debug

# 2. Run the generator script
dotnet script BuildDataModelTemplate.csx

# 3. Commit the updated template
git add TestAssets/DataModelTemplate.xlsx
git commit -m "test: Update Data Model template structure"
```

**That's it!** No need to edit test files.

### Template Requirements

The template contains:
- ✅ 3 tables loaded into Data Model: Sales, Customers, Products
- ✅ 2 relationships: Sales[CustomerID]→Customers[ID], Sales[ProductID]→Products[ID]
- ✅ 3 measures with different format types:
  - Total Sales (Currency)
  - Average Sale (Decimal)
  - Total Customers (WholeNumber)

## Performance Expectations

| Test Type | Setup Time | Per-Test Execution |
|-----------|------------|-------------------|
| **READ tests** (with template) | ~0.5s | ~1-2s |
| **WRITE tests** (fresh file) | ~60-120s | ~2-5s |

**Overall improvement**: ~60% faster test suite when template is used.

## Maintenance

### When to Regenerate Template

Regenerate the template **only when** you need to change:
1. Table structure (add/remove columns)
2. Relationships (add/remove/modify)
3. Measures (add/remove/modify DAX)

### How to Regenerate

See "Generating the Template" section above for complete instructions.

The template is stored in git, so regeneration is rarely needed.

## Troubleshooting

### Template Not Found Error

```
FileNotFoundException: Data Model template not found.
```

**Cause**: Template file is missing from the repository.

**Solution**: 
1. Check if the file exists in git: `git ls-files | grep DataModelTemplate.xlsx`
2. If missing, restore from git history or regenerate (see above)

### Tests Still Slow

**Cause**: Tests might not be using `DataModelReadTestsFixture`.

**Solution**: 
- READ tests: Use `IClassFixture<DataModelReadTestsFixture>`
- WRITE tests: Use `IClassFixture<DataModelWriteTestsFixture>`

### Template Locked by Excel

**Cause**: Excel has the template file open.

**Solution**: Close Excel or kill the process:
```powershell
taskkill /F /IM EXCEL.EXE
```

## Files

| File | Purpose |
|------|---------|
| `Helpers/DataModelReadTestsFixture.cs` | Copies template for READ tests |
| `Helpers/DataModelWriteTestsFixture.cs` | Creates fresh file for WRITE tests |
| `TestAssets/DataModelTemplate.xlsx` | Pre-built template (stored in git) |
| `TestAssets/CreateDataModelAsset.cs` | Builder for regenerating template (rarely used) |

## Summary

- ✅ Template is stored in git - just use it
- ✅ Regeneration is rarely needed (only for structural changes)
- ✅ READ tests are fast (~0.5s setup via template copy)
- ✅ WRITE tests create fresh files (~60-120s setup)
