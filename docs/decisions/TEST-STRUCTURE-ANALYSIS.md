# Core.Tests Structure Analysis - Issues Found

## 🚨 Critical Issues

### 1. Directory Structure Mismatch

**Core Commands (src/ExcelMcp.Core/Commands):**
- Connection/
- DataModel/
- **NamedRange/** ← Core uses this name
- PivotTable/
- PowerQuery/
- Range/
- Sheet/
- Table/
- Vba/
- FileCommands.cs (root level)

**Test Commands (tests/ExcelMcp.Core.Tests/Integration/Commands):**
- Connection/
- DataModel/
- **Parameter/** ← Tests use different name! Should be "NamedRange"
- PivotTable/
- PowerQuery/
- Range/
- **Script/** ← Tests use different name! Should be "Vba"
- Sheet/
- Table/
- **VbaTrust/** ← Separate from Vba tests
- **FileOperations/** ← Should be "File" to match FileCommands

### 2. Namespace Inconsistencies

**Issue:** Namespaces don't follow consistent pattern

| Directory | Namespace | Should Be |
|-----------|-----------|-----------|
| Range/ | `Sbroenne.ExcelMcp.Core.Tests.Integration.Range` | `Sbroenne.ExcelMcp.Core.Tests.Commands.Range` |
| PowerQuery/ (one file) | `Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.PowerQuery` | `Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery` |
| PowerQuery/ (others) | `Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery` | ✅ Correct |
| FileOperations/ | `Sbroenne.ExcelMcp.Core.Tests.Commands.FileOperations` | Should be `.Commands.File` |
| Parameter/ | `Sbroenne.ExcelMcp.Core.Tests.Commands.Parameter` | Should be `.Commands.NamedRange` |
| Script/ | `Sbroenne.ExcelMcp.Core.Tests.Commands.Script` | Should be `.Commands.Vba` |

### 3. Test Class Name Issues

**Separate test classes that should be partial:**

| File | Class Name | Issue |
|------|------------|-------|
| Sheet/SheetTabColorTests.cs | `SheetTabColorTests` | Should be `SheetCommandsTests` (partial) |
| Sheet/SheetVisibilityTests.cs | `SheetVisibilityTests` | Should be `SheetCommandsTests` (partial) |
| Parameter/ParameterCommandsTests.cs | `ParameterCommandsTests` | Should be `NamedRangeCommandsTests` |
| Script/ScriptCommandsTests.cs | `ScriptCommandsTests` | Should be `VbaCommandsTests` |
| VbaTrust/VbaTrustDetectionTests.cs | `VbaTrustDetectionTests` | Should be `VbaCommandsTests` (partial) |
| FileOperations/FileCommandsTests.cs | `FileCommandsTests` | ✅ Correct name, wrong directory |

### 4. Missing Test Directories

**Core commands without test coverage:**
- ✅ All major commands have tests, but they're mislabeled

## 📋 Required Reorganization

### Phase 1: Rename Directories to Match Core

```bash
# Rename mismatched directories
tests/Integration/Commands/Parameter/      → tests/Integration/Commands/NamedRange/
tests/Integration/Commands/Script/         → tests/Integration/Commands/Vba/
tests/Integration/Commands/FileOperations/ → tests/Integration/Commands/File/
tests/Integration/Commands/VbaTrust/       → Merge into tests/Integration/Commands/Vba/
```

### Phase 2: Fix Namespaces

All test namespaces should follow pattern:
```csharp
namespace Sbroenne.ExcelMcp.Core.Tests.Commands.<FeatureName>;
```

**Files to fix:**
- `Range/*.cs` - Change from `.Integration.Range` to `.Commands.Range`
- `PowerQuery/PowerQuerySuccessErrorRegressionTests.cs` - Change from `.Integration.Commands.PowerQuery` to `.Commands.PowerQuery`

### Phase 3: Consolidate Test Classes

**Sheet Tests:**
- `SheetTabColorTests.cs` → Rename to `SheetCommandsTests.TabColor.cs`, make partial
- `SheetVisibilityTests.cs` → Rename to `SheetCommandsTests.Visibility.cs`, make partial

**Vba Tests:**
- `Script/ScriptCommandsTests.cs` → Move to `Vba/VbaCommandsTests.cs`
- `VbaTrust/VbaTrustDetectionTests*.cs` → Rename to `Vba/VbaCommandsTests.Trust*.cs`, make partial

**Named Range Tests:**
- `Parameter/ParameterCommandsTests*.cs` → Rename to `NamedRange/NamedRangeCommandsTests*.cs`

### Phase 4: Fix Test Method Names

Current naming patterns vary:
- ✅ `SetTabColor_WithValidRGB_SetsColorCorrectly` (good)
- ✅ `ShowAsync_MakesHiddenSheetVisible` (good)  
- Need to audit all test names for consistency

## 🎯 Recommended Structure

```
tests/ExcelMcp.Core.Tests/Integration/Commands/
├── Connection/
│   ├── ConnectionCommandsTests.cs
│   ├── ConnectionCommandsTests.List.cs
│   └── ConnectionCommandsTests.View.cs
├── DataModel/
│   ├── DataModelCommandsTests.cs
│   ├── DataModelCommandsTests.Discovery.cs
│   ├── DataModelCommandsTests.Measures.cs
│   ├── DataModelCommandsTests.Relationships.cs
│   └── DataModelCommandsTests.Tables.cs
├── File/                              ← RENAME from FileOperations
│   ├── FileCommandsTests.cs
│   ├── FileCommandsTests.CreateEmpty.cs
│   └── FileCommandsTests.TestFile.cs
├── NamedRange/                        ← RENAME from Parameter
│   ├── NamedRangeCommandsTests.cs     ← RENAME from ParameterCommandsTests
│   ├── NamedRangeCommandsTests.Lifecycle.cs
│   └── NamedRangeCommandsTests.Values.cs
├── PivotTable/
│   ├── PivotTableCommandsTests.cs
│   └── PivotTableCommandsTests.Creation.cs
├── PowerQuery/
│   ├── PowerQueryCommandsTests.cs
│   ├── PowerQueryCommandsTests.Advanced.cs
│   ├── PowerQueryCommandsTests.Lifecycle.cs
│   ├── PowerQueryCommandsTests.LoadConfig.cs
│   ├── PowerQueryCommandsTests.Refresh.cs
│   └── PowerQuerySuccessErrorRegressionTests.cs
├── Range/
│   ├── RangeCommandsTests.cs
│   ├── RangeCommandsTests.Discovery.cs
│   ├── RangeCommandsTests.Editing.cs
│   ├── RangeCommandsTests.Formulas.cs
│   ├── RangeCommandsTests.Hyperlinks.cs
│   ├── RangeCommandsTests.NamedRanges.cs  ← Might belong in NamedRange/
│   ├── RangeCommandsTests.NumberFormat.cs
│   ├── RangeCommandsTests.Search.cs
│   └── RangeCommandsTests.Values.cs
├── Sheet/
│   ├── SheetCommandsTests.cs
│   ├── SheetCommandsTests.Lifecycle.cs
│   ├── SheetCommandsTests.TabColor.cs     ← RENAME from SheetTabColorTests.cs
│   └── SheetCommandsTests.Visibility.cs   ← RENAME from SheetVisibilityTests.cs
├── Table/
│   ├── TableCommandsTests.cs
│   ├── TableCommandsTests.Lifecycle.cs
│   └── TableCommandsTests.StructuredReferences.cs
└── Vba/                               ← RENAME from Script, MERGE VbaTrust
    ├── VbaCommandsTests.cs            ← RENAME from ScriptCommandsTests.cs
    ├── VbaCommandsTests.Trust.cs      ← RENAME from VbaTrustDetectionTests.cs
    ├── VbaCommandsTests.TrustScope.cs ← RENAME from VbaTrustDetectionTests.TrustScope.cs
    └── VbaCommandsTests.ScriptCommands.cs ← RENAME from VbaTrustDetectionTests.ScriptCommands.cs
```

## 🔍 Specific Examples of Test Name Issues

### Example 1: Range Tests - Misplaced NamedRange Test
```
Range/RangeCommandsTests.NamedRanges.cs
```
**Issue:** This tests named ranges, not range operations  
**Fix:** Move to `NamedRange/NamedRangeCommandsTests.RangeOperations.cs`

### Example 2: Sheet Tests - Separate Classes
```
Sheet/SheetTabColorTests.cs         - Separate class
Sheet/SheetVisibilityTests.cs       - Separate class
Sheet/SheetCommandsTests.cs         - Main class
Sheet/SheetCommandsTests.Lifecycle.cs - Partial class
```
**Issue:** Inconsistent - some features are partials, some are separate  
**Fix:** Make ALL into partials of `SheetCommandsTests`

### Example 3: Namespace Mismatch
```csharp
// Range/RangeCommandsTests.cs
namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Range;  // ❌ Inconsistent

// Other tests
namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;    // ✅ Correct pattern
```
**Fix:** All should use `.Commands.<Feature>` pattern

## 📊 Impact Assessment

### Files to Rename: 15+
- 3 directories (Parameter→NamedRange, Script→Vba, FileOperations→File)
- 2 test files for Sheet (TabColor, Visibility)
- All Parameter test files
- All Script test files  
- All VbaTrust test files

### Namespaces to Fix: 10+ files
- All Range test files (8 files)
- PowerQuerySuccessErrorRegressionTests (1 file)
- After directory renames, update namespaces accordingly

### Test Class Names to Fix: 10+ files
- ParameterCommandsTests → NamedRangeCommandsTests
- ScriptCommandsTests → VbaCommandsTests
- SheetTabColorTests → SheetCommandsTests (partial)
- SheetVisibilityTests → SheetCommandsTests (partial)
- VbaTrustDetectionTests → VbaCommandsTests (partial)

## 🚀 Benefits of Reorganization

1. **Consistency** - Directory names match Core commands exactly
2. **Discoverability** - Easy to find tests for any command
3. **Maintainability** - Clear 1:1 mapping between Core and Tests
4. **Navigation** - IDE navigation works better with consistent naming
5. **Onboarding** - New developers can understand structure instantly

## ⚠️ Migration Considerations

### Keep Tests Passing During Migration
- Rename directories one at a time
- Update namespaces immediately after rename
- Run tests after each change
- Use git mv to preserve history

### Order of Operations
1. Create mapping document (this file)
2. Rename directories (preserves structure)
3. Fix namespaces (build will fail until fixed)
4. Rename test files (cosmetic, but important)
5. Update class names (partial classes)
6. Verify all tests pass
7. Update documentation

## 📝 Checklist

### Directory Renames
- [ ] Parameter/ → NamedRange/
- [ ] Script/ → Vba/
- [ ] FileOperations/ → File/
- [ ] VbaTrust/ → Merge into Vba/

### Namespace Fixes
- [ ] All Range/*.cs files
- [ ] PowerQuerySuccessErrorRegressionTests.cs
- [ ] After directory renames, update all affected files

### Class Renames
- [ ] ParameterCommandsTests → NamedRangeCommandsTests
- [ ] ScriptCommandsTests → VbaCommandsTests
- [ ] SheetTabColorTests → SheetCommandsTests (partial)
- [ ] SheetVisibilityTests → SheetCommandsTests (partial)
- [ ] VbaTrustDetectionTests → VbaCommandsTests (partial)

### File Renames
- [ ] SheetTabColorTests.cs → SheetCommandsTests.TabColor.cs
- [ ] SheetVisibilityTests.cs → SheetCommandsTests.Visibility.cs
- [ ] All Parameter test files (prefix with NamedRange)
- [ ] All Script test files (prefix with Vba)
- [ ] All VbaTrust files (prefix with VbaCommands.Trust)

### Documentation Updates
- [ ] Update test documentation
- [ ] Update contributing guide
- [ ] Update architecture documentation
