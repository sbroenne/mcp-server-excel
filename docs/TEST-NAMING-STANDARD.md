# Test Naming Standard - ExcelMcp

## Overview

All integration tests MUST follow this naming convention for consistency, readability, and maintainability.

## Standard Format

```
MethodName_StateUnderTest_ExpectedBehavior
```

### Components

1. **MethodName** - The Core command method being tested (without "Async" suffix)
2. **StateUnderTest** - The scenario, condition, or input state
3. **ExpectedBehavior** - What should happen (success/error/specific result)

### Rules

1. **PascalCase for each component** (no underscores within components)
2. **Separate components with single underscore** (`_`)
3. **No "Async" suffix** on method name (implied)
4. **Descriptive state** (not just "Valid" - be specific)
5. **Clear expected behavior** (Returns*, Creates*, Throws*, etc.)

## Good Examples

### CORRECT Naming

```csharp
// CRUD Operations
[Fact]
public async Task List_EmptyWorkbook_ReturnsEmptyList()

[Fact]
public async Task Create_ValidName_ReturnsSuccess()

[Fact]
public async Task Delete_ExistingItem_ReturnsSuccess()

[Fact]
public async Task Update_NonExistentItem_ReturnsNotFoundError()

// Complex Scenarios
[Fact]
public async Task Import_QueryReferencingAnotherQuery_LoadsDataSuccessfully()

[Fact]
public async Task Refresh_WithConnectionError_ReturnsErrorMessage()

[Fact]
public async Task CreateFromRange_ValidDataWithHeaders_CreatesPivotTable()

// Edge Cases
[Fact]
public async Task Create_NameExceedsMaxLength_ReturnsTruncatedName()

[Fact]
public async Task Delete_LastRemainingSheet_ReturnsError()

// Persistence/Round-trip
[Fact]
public async Task Create_NewTable_PersistsAfterReopen()
```

### INCORRECT Naming

```csharp
// Too vague
[Fact]
public async Task List_WithValidFile_ReturnsSuccess() // "Valid" is not descriptive

// Missing expected behavior
[Fact]
public async Task Create_WithValidName() // What happens?

// Inconsistent separator
[Fact]
public async Task ListAfterImport() // Missing underscores

// Too generic
[Fact]
public async Task TestCreate() // No state or expectation

// Method suffix included
[Fact]
public async Task ListAsync_EmptyWorkbook_ReturnsEmptyList() // Remove "Async"
```

## Pattern Catalog

### Basic CRUD Patterns

| Operation | Pattern | Example |
|-----------|---------|---------|
| **List** | `List_<State>_Returns<Result>` | `List_EmptyWorkbook_ReturnsEmptyList` |
| **Create** | `Create_<Input>_<Outcome>` | `Create_ValidName_ReturnsSuccess` |
| **View/Get** | `View_<Item>_Returns<Data>` | `View_ExistingTable_ReturnsMetadata` |
| **Update** | `Update_<Item>_<Outcome>` | `Update_ExistingQuery_ReturnsSuccess` |
| **Delete** | `Delete_<Item>_<Outcome>` | `Delete_NonExistentItem_ReturnsError` |

### Error Case Patterns

| Scenario | Pattern | Example |
|----------|---------|---------|
| **Not Found** | `<Method>_NonExistent<Item>_ReturnsNotFoundError` | `View_NonExistentQuery_ReturnsNotFoundError` |
| **Invalid Input** | `<Method>_Invalid<Input>_ReturnsValidationError` | `Create_InvalidName_ReturnsValidationError` |
| **Permission** | `<Method>_WithoutPermission_ReturnsUnauthorizedError` | `Delete_WithoutPermission_ReturnsUnauthorizedError` |
| **Conflict** | `<Method>_DuplicateName_ReturnsConflictError` | `Create_DuplicateName_ReturnsConflictError` |

### Integration Patterns

| Scenario | Pattern | Example |
|----------|---------|---------|
| **Persistence** | `<Method>_<Item>_PersistsAfterReopen` | `Create_NewTable_PersistsAfterReopen` |
| **Workflow** | `<Method1>Then<Method2>_<State>_<Outcome>` | `ImportThenDelete_ValidQuery_RemovedFromList` |
| **Side Effects** | `<Method>_<Input>_<SideEffect>` | `Delete_Sheet_UpdatesActiveSheet` |

### State-Specific Patterns

| State | Pattern | Example |
|-------|---------|---------|
| **Empty** | `<Method>_EmptyWorkbook_<Outcome>` | `List_EmptyWorkbook_ReturnsEmptyList` |
| **With Data** | `<Method>_WorkbookWithData_<Outcome>` | `Refresh_WorkbookWithData_UpdatesValues` |
| **Multiple Items** | `<Method>_MultipleItems_<Outcome>` | `Delete_WorkbookWithMultipleSheets_RemovesTargetOnly` |

## Feature-Specific Guidelines

### DataModel Tests
- **Fixture-based**: `<Method>_<FixtureState>_<Outcome>`
- Example: `ListMeasures_RealisticDataModel_ReturnsMeasuresWithFormulas`

### PowerQuery Tests
- **Fixture-based**: `<Method>_<FixtureQuery>_<Outcome>` or `<Method>_<UniqueQuery>_<Outcome>`
- Example: `View_BasicQuery_ReturnsMCode`
- Example: `Import_NewQuery_AddsToWorkbook`

### VBA Tests
- **Trust required**: `<Method>_WithTrustEnabled_<Outcome>`
- Example: `Run_WithTrustEnabled_ExecutesMacro`

### Range Tests
- **Clear suffix**: `Clear<Variant>_<State>_<Outcome>`
- Example: `ClearContents_FormattedRange_PreservesFormatting`
- Example: `ClearAll_UsedRange_RemovesEverything`

### PivotTable Tests
- **Creation source**: `CreateFrom<Source>_<Input>_<Outcome>`
- Example: `CreateFromRange_ValidDataWithHeaders_CreatesPivotTable`
- Example: `CreateFromTable_ExistingTable_CreatesPivotTable`

## Validation Checklist

Before committing test code, verify each test name:

- [ ] Follows `MethodName_StateUnderTest_ExpectedBehavior` pattern
- [ ] Uses PascalCase within components
- [ ] Separates components with single underscore
- [ ] Omits "Async" suffix from method name
- [ ] State is descriptive (not generic like "Valid")
- [ ] Expected behavior is clear and actionable
- [ ] Matches pattern from catalog above
- [ ] Test body matches what the name promises

## Common Mistakes to Fix

### 1. Too Generic States

```csharp
// BEFORE
List_WithValidFile_ReturnsSuccessResult

// AFTER
List_EmptyWorkbook_ReturnsEmptyList
List_WorkbookWithQueries_ReturnsList
```

### 2. Missing Expected Behavior

```csharp
// BEFORE
Create_WithValidParameter

// AFTER
Create_ValidParameter_ReturnsSuccess
```

### 3. Redundant Suffixes

```csharp
// BEFORE
Import_WithValidMCode_ReturnsSuccessResult

// AFTER
Import_ValidMCode_ReturnsSuccess
```

### 4. Method Suffix Included

```csharp
// BEFORE
ListAsync_EmptyWorkbook_ReturnsEmptyList

// AFTER
List_EmptyWorkbook_ReturnsEmptyList
```

### 5. Workflow Tests Not Clear

```csharp
// BEFORE
Import_ThenDelete_ThenList_ShowsEmpty

// AFTER
ImportThenDelete_ValidQuery_RemovedFromList
```

## Migration Strategy

### Phase 1: Document Current State 
- [x] Analyze all 130+ test names
- [x] Identify patterns and inconsistencies
- [x] Create naming standard

### Phase 2: Systematic Renaming
1. Generate rename script for each feature
2. Review and approve renames
3. Execute renames (preserves git history)
4. Verify build and tests

### Phase 3: Enforce Standard
- Add to PR review checklist
- Update copilot instructions
- Include in CONTRIBUTING.md

## References

### xUnit Best Practices
- [Microsoft: Unit test naming](https://docs.microsoft.com/en-us/dotnet/core/testing/unit-testing-best-practices#best-practices)
- [Roy Osherove: Naming standards](http://osherove.com/blog/2005/4/3/naming-standards-for-unit-tests.html)
- [Vladimir Khorikov: Test naming](https://enterprisecraftsmanship.com/posts/you-naming-tests-wrong/)

### Pattern: AAA (Arrange-Act-Assert)
Test names should reflect what is being tested, under what conditions, and what the expected outcome is.

```csharp
[Fact]
public async Task MethodName_StateUnderTest_ExpectedBehavior()
{
 // Arrange - Set up test state
 
 // Act - Execute the operation
 
 // Assert - Verify expected behavior
}
```
