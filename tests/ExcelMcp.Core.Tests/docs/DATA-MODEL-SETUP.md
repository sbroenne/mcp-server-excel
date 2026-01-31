# Data Model Test Setup

> **Fixture-as-Test pattern for Data Model integration tests**

## Overview

Data Model tests use a **fixture-as-test pattern** where the fixture initialization IS the test for data model creation:
- **Fixture initialization**: ~60-120 seconds (creates data model and validates all creation commands)
- **Per-test execution**: ~1-2 seconds (uses shared data model file)
- **Performance gain**: 95% faster than creating data model per test

## Architecture

### Single Unified Fixture

**DataModelTestsFixture** (`Helpers/DataModelTestsFixture.cs`)
- Creates ONE Data Model file per test CLASS during initialization
- Fixture initialization IS the test - validates all creation commands:
  - `ExcelBatch.CreateNewWorkbook()` to create new file with session (optimized single start)
  - `TableCommands.AddToDataModelAsync()` for all tables
  - `DataModelCommands.CreateRelationshipAsync()` for all relationships
  - `DataModelCommands.CreateMeasureAsync()` for all measures
  - `Batch.SaveAsync()` persistence
- Each test gets its own batch/session (isolation at batch level)
- Write operations use unique names to avoid conflicts
- Exposes `CreationResult` for validation tests

### Data Model Structure

**Created by fixture:**
- 3 Excel Tables: SalesTable (10 rows), CustomersTable (5 rows), ProductsTable (5 rows)
- 2 Relationships: SalesTable→CustomersTable, SalesTable→ProductsTable
- 3 DAX Measures: Total Sales, Average Sale, Total Customers

## Using the Fixture

### Standard Test Pattern

```csharp
[Trait("Category", "Integration")]
[Trait("Feature", "DataModel")]
public partial class DataModelCommandsTests : IClassFixture<DataModelTestsFixture>
{
    private readonly DataModelCommands _commands;
    private readonly string _dataModelFile;
    private readonly DataModelCreationResult _creationResult;

    public DataModelCommandsTests(DataModelTestsFixture fixture)
    {
        _commands = new DataModelCommands();
        _dataModelFile = fixture.TestFilePath;
        _creationResult = fixture.CreationResult;
    }

    [Fact]
    public async Task ListTables_ReturnsExpectedTables()
    {
        // Each test gets its own batch (isolated session)
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        var result = await _commands.ListTablesAsync(batch);
        
        Assert.True(result.Success);
        Assert.Equal(3, result.Tables.Count);
    }
}
```

### Write Operations Pattern

```csharp
[Fact]
public async Task CreateMeasure_WithValidParameters_CreatesSuccessfully()
{
    // Use unique name to avoid conflicts with other tests
    var measureName = $"Test_{nameof(CreateMeasure_WithValidParameters_CreatesSuccessfully)}_{Guid.NewGuid():N}";
    
    await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
    var result = await _commands.CreateMeasureAsync(
        batch, "SalesTable", measureName, "SUM(SalesTable[Amount])");
    
    Assert.True(result.Success);
    
    // Verify it exists
    var listResult = await _commands.ListMeasuresAsync(batch);
    Assert.Contains(listResult.Measures, m => m.Name == measureName);
}
```

### Creation Validation Test

```csharp
[Fact]
public void DataModelCreation_ViaFixture_CreatesCompleteModel()
{
    // Assert the fixture creation succeeded
    Assert.True(_creationResult.Success, 
        $"Data Model creation failed: {_creationResult.ErrorMessage}");
    Assert.True(_creationResult.FileCreated);
    Assert.Equal(3, _creationResult.TablesCreated);
    Assert.Equal(3, _creationResult.TablesLoadedToModel);
    Assert.Equal(2, _creationResult.RelationshipsCreated);
    Assert.Equal(3, _creationResult.MeasuresCreated);
    
    // This test appears in results as proof creation was tested
}
```

## Key Benefits

### ✅ Fixture IS the Creation Test
- Validates FileCommands.CreateEmptyAsync()
- Validates TableCommands.AddToDataModelAsync()
- Validates DataModelCommands.CreateRelationshipAsync()
- Validates DataModelCommands.CreateMeasureAsync()
- Validates Batch.SaveAsync() persistence
- If creation fails, all tests fail (correct - no point testing if foundation broken)

### ✅ Fast Tests
- 60-120s setup ONCE per test class
- ~1-2s per test execution
- 95% faster than per-test creation

### ✅ Test Isolation
- Each test gets its own batch/session
- Write tests use unique names (no cross-contamination)
- No file sharing between test classes

### ✅ Transparent
- Setup code visible and maintainable
- No binary template files
- Uses production commands (proves they work!)

### ✅ Visible in Test Results
```
✅ DataModelCreation_ViaFixture_CreatesCompleteModel (0.1s)
✅ DataModelCreation_Persists_AfterReopenFile (1.2s)
✅ ListTables_ReturnsExpectedTables (1.1s)
✅ CreateMeasure_WithValidParameters_CreatesSuccessfully (2.3s)
```

## Performance Expectations

| Operation | Time | Notes |
|-----------|------|-------|
| **Fixture initialization** | 60-120s | Creates data model, runs ONCE per class |
| **Per-test execution** | 1-2s | Uses shared file with own batch |
| **Overall improvement** | 95% | vs creating data model per test |

## Test Organization

All data model tests are in one test class with partial files:
- `DataModelCommandsTests.cs` - Base class with fixture, creation validation tests
- `DataModelCommandsTests.Discovery.cs` - List/View/Get operations
- `DataModelCommandsTests.Tables.cs` - Table operations
- `DataModelCommandsTests.Measures.cs` - Measure CRUD (uses unique names)
- `DataModelCommandsTests.Relationships.cs` - Relationship CRUD

## Fixture Output

When tests run, you'll see:
```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TESTING: Data Model Creation (via fixture initialization)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  [1/6] Testing: File creation...
        ✅ File created successfully
  [2/6] Testing: Table creation (3 tables with data)...
        ✅ Created 3 tables: SalesTable, CustomersTable, ProductsTable
  [3/6] Testing: TableCommands.AddToDataModelAsync() for 3 tables...
        ✅ All 3 tables loaded into Data Model
  [4/6] Testing: DataModelCommands.CreateRelationshipAsync() for 2 relationships...
        ✅ Created 2 relationships: Sales→Customers, Sales→Products
  [5/6] Testing: DataModelCommands.CreateMeasureAsync() for 3 measures...
        ✅ Created 3 measures: Total Sales, Average Sale, Total Customers
  [6/6] Testing: Batch.SaveAsync() persistence...
        ✅ Data Model saved successfully

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✅ CREATION TEST PASSED in 87.3s
   📊 3 tables created and loaded
   🔗 2 relationships established
   📏 3 DAX measures defined
   💾 File: C:\Temp\DataModelTests_abc123\DataModel.xlsx
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

## Troubleshooting

### All Tests Fail With "Creation Failed"

**Cause**: Fixture initialization failed during data model creation.

**Solution**: 
1. Check fixture output for specific error
2. Verify Excel is installed and accessible
3. Check TOM library availability (Data Model requires Excel 2013+)

### Tests Interfere With Each Other

**Cause**: Write tests not using unique names.

**Solution**: Use pattern `$"Test_{nameof(TestMethod)}_{Guid.NewGuid():N}"`

### Slow Test Execution

**Cause**: Tests might not be using shared file from fixture.

**Solution**: Ensure all tests use `_dataModelFile` from fixture, not creating new files.

## Files

| File | Purpose |
|------|---------|
| `Helpers/DataModelTestsFixture.cs` | Unified fixture (creates data model and validates creation) |
| `Integration/Commands/DataModel/DataModelCommandsTests.cs` | Base class with creation validation |
| `Integration/Commands/DataModel/DataModelCommandsTests.*.cs` | Partial classes for different operations |

## Summary

- ✅ **Fixture IS the creation test** - validates all creation commands during initialization
- ✅ **No template files** - data model built from production code
- ✅ **Fast tests** - 60-120s setup once, then 1-2s per test
- ✅ **Test isolation** - each test gets own batch, write tests use unique names
- ✅ **Transparent** - all creation code visible and maintainable
- ✅ **Fail-fast** - if creation fails, all tests fail (correct behavior)
