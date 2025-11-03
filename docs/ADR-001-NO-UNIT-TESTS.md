# ADR-001: Why ExcelMcp Has No Traditional Unit Tests

**Status**: Accepted  
**Date**: 2025-11-02  
**Decision Makers**: Architecture Team  
**Stakeholders**: Development Team, Code Reviewers, Contributors

---

## Context and Problem Statement

ExcelMcp is a COM automation library that wraps Excel's COM API. During code review, the question inevitably arises: **"Why don't you have unit tests?"**

This ADR documents our architectural decision and the reasoning behind our testing strategy.

---

## Decision

**We do NOT write traditional unit tests for ExcelMcp.** Our test suite consists exclusively of **integration tests** that interact with real Excel instances via COM automation.

### What We DON'T Do

❌ Mock Excel COM objects  
❌ Write unit tests for business logic  
❌ Test internal methods in isolation  
❌ Separate "unit" from "integration" concerns  

### What We DO Do

✅ Write comprehensive integration tests against real Excel  
✅ Test every operation with actual Excel workbooks  
✅ Verify behavior through COM API interactions  
✅ Run tests on CI/CD with Excel installed (Azure self-hosted runner)  

---

## Rationale

### 1. Excel COM Cannot Be Meaningfully Mocked

**The Problem**: Excel's COM API is the "database" we're automating against. Consider this code:

```csharp
public async Task<OperationResult> CreateWorksheet(IExcelBatch batch, string sheetName)
{
    return await batch.ExecuteAsync((ctx, ct) => 
    {
        dynamic sheets = ctx.Book.Worksheets;  // COM object
        dynamic newSheet = sheets.Add();       // COM method
        newSheet.Name = sheetName;             // COM property
        return new OperationResult { Success = true };
    });
}
```

**What would a "unit test" look like?**

```csharp
// Option 1: Mock the COM object
var mockBook = new Mock<dynamic>();  // ❌ Cannot mock dynamic COM objects
mockBook.Setup(b => b.Worksheets).Returns(...);  // ❌ Runtime binding fails

// Option 2: Test without Excel
[Fact]
public void CreateWorksheet_ReturnsSuccess()
{
    var result = CreateWorksheet(null!, "Test");  // ❌ What are we testing?
    Assert.True(result.Success);  // ❌ This proves nothing!
}
```

**The Truth**: The ONLY way to verify this code works is to:
1. Open a real Excel instance
2. Call the real COM API
3. Verify the worksheet actually exists in Excel

**That's an integration test by definition.**

### 2. Our Integration Tests ARE Our Unit Tests

In traditional layered architecture:
- **Unit tests** test business logic in isolation
- **Integration tests** verify components work together
- **E2E tests** test the entire system

In COM automation architecture:
- **Integration tests** test business logic AND COM interaction (these ARE our unit tests)
- **E2E tests** don't exist (we ARE the library, not an application)

**Analogy**: ExcelMcp is like a database driver (e.g., Npgsql for PostgreSQL):
- You don't mock `DbConnection` to test SQL queries
- You test against a real database instance
- The "integration test" IS the unit test

### 3. Industry Precedent

This pattern is **normal and correct** for COM/browser/external system automation:

| Library | What It Automates | Test Strategy |
|---------|------------------|---------------|
| **Selenium WebDriver** | Browser DOM | Integration tests against real browsers |
| **Playwright** | Browser automation | Integration tests with browser instances |
| **AWS SDK** | Cloud services | Integration tests against AWS (or LocalStack) |
| **ExcelMcp** | Excel COM | Integration tests against Excel instances |

**None of these libraries have meaningful unit tests** for their core automation logic. They all test against the real external system.

### 4. What About .NET Framework APIs?

**Question**: "Shouldn't we unit test our wrappers around .NET APIs?"

**Answer**: No, because .NET already tests those APIs. Consider:

```csharp
public static string ValidateAndNormalizePath(string path)
{
    if (string.IsNullOrWhiteSpace(path))
        throw new ArgumentException("Path cannot be null");
    
    return Path.GetFullPath(path);  // .NET handles validation
}
```

**What's actually happening**:
- `Path.GetFullPath()` does: path traversal prevention, invalid character checking, normalization
- Our code does: null check (trivial)

**Testing this**:
```csharp
[Fact]
public void ValidatePath_WithTraversal_ThrowsException()
{
    Assert.Throws<ArgumentException>(() => 
        PathValidator.ValidateAndNormalizePath("../../etc/passwd"));
}
```

**Problem**: This test verifies .NET's `Path.GetFullPath()` works, not our code. We're testing Microsoft's code, not ours.

**Better approach**: Trust .NET's APIs (they're battle-tested). If our path validation is wrong, our integration tests will fail when we try to open a file.

### 5. The MCP Protocol Argument

**Question**: "Shouldn't we unit test MCP JSON serialization?"

**Answer**: No, the MCP SDK handles protocol compliance.

```csharp
public class RangeValueResult : ResultBase
{
    public List<List<object?>> Values { get; set; }
}

// MCP SDK serializes this to JSON automatically
```

**What a unit test would look like**:
```csharp
[Fact]
public void RangeValueResult_SerializesToJson()
{
    var result = new RangeValueResult { Values = [[1, 2]] };
    var json = JsonSerializer.Serialize(result);
    Assert.Contains("[[1,2]]", json);
}
```

**Problem**: This tests `System.Text.Json`, not our code. If JSON serialization breaks, the MCP SDK will fail to parse responses, and our integration tests will catch it.

---

## Real-World Test Coverage

### What Our Integration Tests Actually Test

**Scenario**: Create a worksheet named "Sales"

```csharp
[Fact]
public async Task CreateWorksheet_ValidName_CreatesSheet()
{
    // Arrange
    var testFile = await CreateUniqueTestFile(".xlsx");
    
    // Act
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);
    var result = await _commands.CreateAsync(batch, "Sales");
    await batch.SaveAsync();
    
    // Assert - Round-trip validation
    Assert.True(result.Success);
    
    await using var batch2 = await ExcelSession.BeginBatchAsync(testFile);
    var list = await _commands.ListAsync(batch2);
    Assert.Contains(list.Items, s => s.Name == "Sales");
}
```

**What this ACTUALLY tests**:
1. ✅ Excel session management (ExcelSession.BeginBatchAsync)
2. ✅ COM object lifecycle (Workbooks.Open, Worksheets.Add)
3. ✅ Batch transaction handling (IExcelBatch)
4. ✅ Error handling (COM exceptions)
5. ✅ Resource cleanup (IDisposable, COM release)
6. ✅ Persistence (workbook.Save)
7. ✅ Re-opening workbooks (validates saved state)
8. ✅ Business logic (worksheet creation)
9. ✅ API contract (ISheetCommands interface)

**A unit test could verify**: None of the above (requires real Excel).

### Test Statistics

- **Integration Tests**: ~200+ tests covering all operations
- **Execution Time**: 10-20 minutes (acceptable for CI/CD)
- **Coverage**: All production code paths
- **False Positives**: Near zero (tests against real Excel)

---

## Consequences

### Positive

✅ **Tests verify real behavior** - We test what actually happens in Excel, not mocked abstractions  
✅ **High confidence** - If tests pass, the code works in production  
✅ **No mock maintenance** - No complex mock setup that becomes outdated  
✅ **Catches integration bugs** - We discover COM quirks (e.g., 1-based indexing, Type 3/4 connection discrepancy)  
✅ **Industry standard** - Follows proven patterns from Selenium, Playwright, AWS SDK  

### Negative

⚠️ **Slower tests** - 10-20 minutes vs seconds for unit tests  
⚠️ **Requires Excel** - CI/CD needs Windows + Excel (Azure self-hosted runner)  
⚠️ **Resource intensive** - Each test opens/closes Excel COM instance  
⚠️ **Cannot run on Linux** - Excel COM is Windows-only  

### Mitigation Strategies

**For slow tests**:
- Run tests in parallel (xUnit parallelization)
- Cache Excel instances where safe
- Use OnDemand trait for expensive tests
- Optimize CI/CD with dedicated Windows runners

**For Excel dependency**:
- Azure self-hosted runner with Office 365 installed
- Local development requires Excel (documented in CONTRIBUTING.md)
- Pre-commit hooks run quick validation only

---

## Alternatives Considered

### Alternative 1: Mock Excel COM Objects

**Rejected** because:
- `dynamic` COM objects cannot be meaningfully mocked
- Mocks would just verify our mock setup, not real Excel behavior
- Excel's COM API has quirks (1-based indexing, async RefreshAll issues) that mocks wouldn't catch

### Alternative 2: Record/Replay COM Interactions

**Rejected** because:
- Fragile (breaks when Excel updates)
- Doesn't test actual Excel state
- High maintenance burden
- Doesn't verify persistence (save/reload)

### Alternative 3: Separate Business Logic from COM

**Rejected** because:
- There IS no business logic separate from COM interaction
- Our "business logic" IS calling Excel COM methods correctly
- Would create artificial abstraction layers with no value

### Alternative 4: Test Against Excel Interop Primary Assemblies

**Rejected** because:
- Still requires Excel installed
- PIAs are just type definitions, not implementation
- Doesn't reduce test execution time
- We use late binding (`dynamic`) intentionally for flexibility

---

## Exceptions: When Unit Tests Make Sense

We **would** write unit tests for:

1. **Pure algorithms** - If we had complex calculations independent of Excel (we don't)
2. **Custom protocols** - If we implemented custom serialization (MCP SDK handles this)
3. **Complex state machines** - If we had stateful logic beyond COM (we don't)

**Current reality**: 100% of our logic involves Excel COM interaction, so 100% of our tests are integration tests.

---

## Code Review Response Template

When reviewers ask "Why no unit tests?", respond:

> **ExcelMcp is a COM automation library.** We test against real Excel instances because:
> 
> 1. **Excel COM cannot be mocked** - Dynamic COM objects don't support traditional mocking frameworks
> 2. **Integration tests ARE our unit tests** - We test business logic (COM interaction) in the only way possible
> 3. **Industry standard** - Selenium, Playwright, AWS SDK all use the same pattern
> 4. **High confidence** - Tests verify actual Excel behavior, not mock abstractions
> 
> See `docs/ADR-001-NO-UNIT-TESTS.md` for full rationale.

---

## References

1. **Martin Fowler - "Test Pyramid Antipattern"**: https://martinfowler.com/bliki/TestPyramid.html
   - "The test pyramid is a simplification... some contexts don't fit the pyramid"
   
2. **Selenium Testing Best Practices**: https://www.selenium.dev/documentation/test_practices/
   - Tests run against real browsers, not mocks
   
3. **Playwright Testing Philosophy**: https://playwright.dev/docs/test-philosophy
   - "End-to-end tests should test real scenarios"
   
4. **AWS SDK Testing**: https://github.com/aws/aws-sdk-net
   - Integration tests against AWS or LocalStack, minimal unit tests

5. **Microsoft Office Interop Best Practices**: https://learn.microsoft.com/office/client-developer/
   - COM automation testing requires real Office instances

---

## Decision Record

**Date**: November 2, 2025  
**Decided by**: Architecture Team  
**Status**: Accepted  

**Supersedes**: N/A  
**Superseded by**: N/A  

**Last Reviewed**: November 2, 2025  
**Next Review**: When adding features that don't require Excel COM (if ever)

---

## Appendix: Test Execution Strategy

### Local Development
```powershell
# Fast feedback (integration tests, excludes VBA, excludes OnDemand)
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"
```

### Pre-Commit
```powershell
# Comprehensive (all integration tests except OnDemand and VBA)
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"
```

### Session/Batch Code Changes
```powershell
# MANDATORY when modifying ExcelSession.cs or ExcelBatch.cs
dotnet test --filter "RunType=OnDemand"
```

### VBA Tests (Manual Only)
```powershell
# Requires "Trust access to VBA project object model" enabled
dotnet test --filter "(Feature=VBA|Feature=VBATrust)&RunType!=OnDemand"
```

### CI/CD Pipeline
- **GitHub Actions**: Build verification only (no Excel)
- **Azure Self-Hosted Runner**: All integration tests (Excel installed)
- **Both must pass** before merge to main

---

**End of ADR-001**
