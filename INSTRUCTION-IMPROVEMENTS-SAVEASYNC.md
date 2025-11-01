# Testing Strategy Instruction Improvements - SaveAsync Timing

## Summary

Updated `.github/instructions/testing-strategy.instructions.md` to prevent the systematic mistake of calling `await batch.SaveAsync()` in the middle of a test.

---

## Problem

Tests were calling `SaveAsync()` mid-test, which commits and closes the batch transaction, causing subsequent operations to fail:

```csharp
// WRONG Pattern
await using var batch = await ExcelSession.BeginBatchAsync(testFile);

var result1 = await _commands.CreateAsync(batch, "Sheet1");
await batch.SaveAsync();  // ❌ Too early - closes batch!

var result2 = await _commands.RenameAsync(batch, "Sheet1", "NewName");  // FAILS!
```

---

## Solution

Added explicit rules and examples throughout testing strategy instructions:

### 1. Updated Batch API Pattern Checklist

**Added to checklist:**
- ✅ **CRITICAL:** `await batch.SaveAsync()` MUST be called ONLY at the END of the test
- ✅ **NEVER** call `SaveAsync()` in the middle of a test (prevents subsequent operations)
- ✅ **NEVER** call `SaveAsync()` multiple times in a single test
- ✅ Only call `SaveAsync()` if test modifies data that needs persistence verification

### 2. Enhanced Batch API Pattern Section

**Added CRITICAL SaveAsync Rules:**
- ✅ Call SaveAsync ONLY at the END of the test
- ✅ Call SaveAsync ONLY ONCE per test
- ✅ Call SaveAsync ONLY if you need to persist modifications
- ❌ NEVER call SaveAsync in the middle of a test
- ❌ NEVER call SaveAsync multiple times
- ❌ Read-only operations do NOT need SaveAsync

### 3. Added Common Mistake #7

**New section with wrong vs correct examples:**

**❌ WRONG Pattern:**
```csharp
[Fact]
public async Task Test_WrongPattern()
{
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);
    
    var result1 = await _commands.CreateAsync(batch, "Sheet1");
    await batch.SaveAsync();  // ❌ WRONG - too early!
    
    var result2 = await _commands.RenameAsync(batch, "Sheet1", "NewName");  // FAILS!
}
```

**✅ CORRECT Pattern:**
```csharp
[Fact]
public async Task Test_CorrectPattern()
{
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);
    
    // All operations
    var result1 = await _commands.CreateAsync(batch, "Sheet1");
    Assert.True(result1.Success);
    
    var result2 = await _commands.RenameAsync(batch, "Sheet1", "NewName");
    Assert.True(result2.Success);
    
    // Save ONLY at the end
    await batch.SaveAsync();  // ✅ CORRECT - after all operations
}
```

**Why This Matters:**
- SaveAsync commits and closes the batch transaction
- No operations can be performed after SaveAsync
- Tests should verify all operations, THEN save once at the end
- If you need to verify persistence, open a NEW batch session to read back

### 4. Updated Test Template

Modified template to show SaveAsync at the end:

```csharp
[Fact]
public async Task Operation_Scenario_ExpectedResult()
{
    var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(...);
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);

    // Act
    var result = await _commands.OperationAsync(batch, args);

    // Assert
    Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");
    Assert.NotNull(result.Data);
    
    // Save only at the end if modifications made
    await batch.SaveAsync();  // ✅ Added to template
}
```

---

## Impact

**Before:** Tests could accidentally call SaveAsync mid-test, causing:
- Subsequent operations failing mysteriously
- Unclear error messages
- Wasted debugging time

**After:** Clear guidance prevents the mistake:
- Checklist item warns developers
- Common Mistakes section shows wrong vs correct
- Template reinforces correct pattern
- CRITICAL markers draw attention

---

## Files Changed

1. `.github/instructions/testing-strategy.instructions.md`
   - Updated Batch API Pattern checklist (4 new rules)
   - Added CRITICAL SaveAsync Rules section
   - Added Common Mistake #7 with examples
   - Updated test template

---

## Commit

```
docs: add critical SaveAsync timing rules to testing strategy

- Add SaveAsync timing to Batch API Pattern checklist
- Add new Common Mistake #7: Calling SaveAsync mid-test
- Add CRITICAL SaveAsync Rules section in Batch API Pattern
- Update test template to show SaveAsync at end
- Emphasize: SaveAsync ONLY at END of test, ONLY ONCE, prevents subsequent operations

Prevents bug where SaveAsync in middle of test breaks subsequent operations.
```

---

## Future Prevention

These instructions will be automatically applied to all test files (`applyTo: "tests/**/*.cs"`), preventing the mistake from recurring.
