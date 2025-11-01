# Data Model Test Optimization - Implementation Status

## ✅ Completed

1. **Created Infrastructure**
   - `DataModelReadTestsFixture.cs` - Fixture that copies pre-built template
   - `DataModelAssetBuilder.cs` - Updated with version metadata
   - `DataModelAssetBuilderTests.cs` - Tests for template validation
   - Updated `DataModelCommandsTests.cs` - Uses template for READ operations

2. **Template Generation Script**
   - `BuildDataModelTemplate.csx` - Standalone script to generate template
   - Includes version metadata in file properties
   - Creates realistic Data Model with 3 tables, 2 relationships, 3 measures

3. **Documentation**
   - `DATA-MODEL-TEST-OPTIMIZATION.md` - Complete strategy document
   - Instructions for regenerating template
   - Version checking approach

## 🚧 TODO - Complete the Implementation

### Step 1: Generate the Template File

The template generation script exists but needs to be run to completion:

```bash
# Kill any Excel processes
Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue

# Generate template (takes 60-120 seconds)
cd D:\source\mcp-server-excel
dotnet script BuildDataModelTemplate.csx
```

**Expected output:**
```
Creating: tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx
Created workbook
Creating Sales table...
Creating Customers table...
Creating Products table...
Adding to Data Model (this takes 30-90 seconds)...
  Sales: True
  Customers: True
  Products: True
Creating relationships...
Creating measures...
✅ Template created in XX.Xs
```

### Step 2: Verify Template

```bash
# Check template was created
ls -l tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx

# Run verification test
dotnet test --filter "FullyQualifiedName~DataModelTemplate_HasExpectedStructure"
```

Should show:
- File exists (~10-15 KB)
- Contains 3 worksheets (Sales, Customers, Products)
- Data Model has 3 tables
- Data Model has 2 relationships
- Data Model has 3 measures

### Step 3: Test Performance Improvement

```bash
# Run Data Model READ tests with template
dotnet test --filter "Feature=DataModel&FullyQualifiedName~ListTables_WithValidFile"
```

**Expected:**
- **Before**: ~10-15 seconds (builds Data Model from scratch)
- **After**: ~1-2 seconds (copies template file)
- **Improvement**: 80-90% faster

### Step 4: Commit Template to Repo

```bash
# Add template file to git
git add tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx

# Commit with version
git commit -m "test: Add Data Model template (v1.0.0) for fast READ tests

Pre-built Data Model template with:
- 3 tables (Sales, Customers, Products)
- 2 relationships (Sales-Customers, Sales-Products)
- 3 measures (Total Sales, Average Sale, Total Customers)

READ tests now copy this template (~0.5s) instead of building from scratch (60-120s).

Expected improvement: 60% faster Data Model test suite."
```

### Step 5: Clean Up

```bash
# Remove build scripts
Remove-Item BuildDataModelTemplate.csx
Remove-Item tests/ExcelMcp.Core.Tests/BuildAsset.csx -ErrorAction SilentlyContinue
```

## 📊 Expected Results

### Before Optimization
- **ListTables test**: 10-15s (builds Data Model)
- **10 READ tests**: 100-150s total
- **Total test suite**: 160-220s

### After Optimization
- **ListTables test**: 1-2s (copies template)
- **10 READ tests**: 10-20s total
- **Total test suite**: 70-90s

**Overall improvement: 60-70% faster**

## 🔄 Maintenance

### When to Regenerate Template

1. **Schema changes**: Adding/removing tables, relationships, measures
2. **Test requirements change**: Need different data or structure
3. **Quarterly refresh**: Update to current conventions

### How to Regenerate

```bash
# Update CreateDataModelAsset.cs with new schema
# Increment ASSET_VERSION constant

# Regenerate template
dotnet script BuildDataModelTemplate.csx

# Verify
dotnet test --filter "FullyQualifiedName~DataModelTemplate_HasExpectedStructure"

# Commit
git add tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx
git commit -m "test: Regenerate Data Model template (v1.1.0 - added XYZ)"
```

## 🐛 Troubleshooting

### Template Not Found

```
Error: Data Model template not found.
```

**Solution**: Run `dotnet script BuildDataModelTemplate.csx` to generate it.

### Template Has Wrong Version

```
Error: Template version mismatch. Expected v1.0.0
```

**Solution**: Regenerate template with current version.

### Tests Still Slow

Check if tests are using `requiresWritableDataModel = true` (forces rebuild).
READ-only tests should use default `requiresWritableDataModel = false`.

## 📁 Files Modified

- ✅ `tests/ExcelMcp.Core.Tests/Helpers/DataModelReadTestsFixture.cs`
- ✅ `tests/ExcelMcp.Core.Tests/Helpers/DataModelAssetBuilderTests.cs`
- ✅ `tests/ExcelMcp.Core.Tests/TestAssets/CreateDataModelAsset.cs`
- ✅ `tests/ExcelMcp.Core.Tests/Integration/Commands/DataModel/DataModelCommandsTests.cs`
- 🚧 `tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx` (needs generation)

## ✨ Next Steps After Template Generation

1. Run full Data Model test suite to verify performance
2. Document actual performance gains in commit message
3. Update CI/CD to verify template exists
4. Consider adding pre-commit hook to warn if CreateDataModelAsset.cs changes without template update
