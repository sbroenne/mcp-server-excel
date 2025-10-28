---
applyTo: "**"
---

# CRITICAL RULES - MUST FOLLOW

> **⚠️ These are NON-NEGOTIABLE rules for all development work on ExcelMcp**

## Rule 1: No Silent Test Failures

**Tests must NEVER silently skip validation or catch exceptions without failing.**

### ❌ FORBIDDEN
```csharp
catch (Exception) { _output.WriteLine("Skipping validation"); }  // WRONG!
```

### ✅ CORRECT  
```csharp
var result = DoOperation();  // Throws if fails - GOOD!
Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
```

**Why:** Silent failures hide bugs and create false confidence.

---

## Rule 2: No NotImplementedException

**NotImplementedException is NEVER acceptable in any feature.**

### Requirements
- ✅ Full Excel COM interop implementation
- ✅ Real test data (not mocks or empty workbooks)
- ✅ All tests pass with actual Excel operations
- ❌ NO placeholder methods

**Why:** Incomplete implementations waste time and provide zero functionality.

---

## Rule 3: Always Run Pool Cleanup Tests

**When modifying `ExcelInstancePool.cs` or `ExcelHelper.cs` pooling code:**

```bash
# MANDATORY before commit
dotnet test --filter "RunType=OnDemand" --list-tests  # Verify 5 tests
dotnet test --filter "RunType=OnDemand"              # All must pass
```

**Why:** Pool bugs cause Excel.exe process leaks in production. OnDemand tests are the ONLY verification.

**Requirements:**
- ⚠️ Excel installed (local execution only)
- ⚠️ Takes 3-5 minutes
- ⚠️ All 5 tests must pass

---

## Rule 4: Update Instructions After Significant Work

**After completing multi-step tasks, update copilot instructions with:**
- Lessons learned
- Architecture changes  
- Testing insights
- Bug fixes and prevention strategies

**Why:** Future AI sessions benefit from accumulated knowledge.

---

## Rule 5: All Changes Via Pull Requests

**NEVER commit directly to `main` branch.**

### Required Process
1. Create feature branch
2. Make changes with tests
3. Create PR with description
4. Wait for CI/CD + code review
5. Merge when approved

**Why:** Branch protection enforces quality gates and review.

---

## Rule 6: COM API First - No External Dependencies for Native Capabilities

**Everything that CAN be implemented via Excel COM API MUST be implemented via COM API.**

### Requirements
- ✅ Use native Excel COM objects and methods
- ❌ NEVER add external libraries for capabilities Excel already provides
- ❌ NEVER use third-party APIs when Excel COM supports the operation
- ⚠️ Only use external libraries (like TOM) for features Excel COM explicitly doesn't support

### Examples

**✅ CORRECT - Use Excel COM API:**
```csharp
// CREATE measure - Excel COM fully supports this
dynamic measures = table.ModelMeasures;
dynamic newMeasure = measures.Add(
    MeasureName: "TotalSales",
    AssociatedTable: table,
    Formula: "SUM(Sales[Amount])",
    FormatInformation: model.ModelFormatCurrency,
    Description: "Total sales amount"
);
```

**❌ WRONG - Don't use TOM when Excel COM works:**
```csharp
// WRONG - Excel COM already supports measure creation!
// Don't use TOM API for this
var tom = new TomServer();
var measure = new Microsoft.AnalysisServices.Tabular.Measure();
// ... unnecessary complexity
```

### Validation Process
1. **Before adding ANY external library:** Search Microsoft official docs for Excel COM capability
2. **If Excel COM supports it:** Use Excel COM API (no exceptions)
3. **If Excel COM doesn't support it:** Document why, then consider alternatives
4. **Always validate against official Microsoft documentation:** https://learn.microsoft.com/en-us/office/vba/api/overview/excel

### Real Example - DataModelCommands

**Original spec claimed:** "Use TOM API for measure creation" ❌ WRONG

**Microsoft official docs proved:** Excel COM fully supports `ModelMeasures.Add()` ✅ CORRECT

**Lesson:** Always validate specs against Microsoft official documentation before architectural decisions.

**Why This Rule Exists:**
- Simpler code (native operations, no external dependencies)
- Better performance (direct COM access vs library overhead)
- Fewer deployment issues (no NuGet packages, no version conflicts)
- Works offline (no server dependencies)
- Smaller attack surface (fewer dependencies = fewer vulnerabilities)

---

## Quick Reference

| Scenario | Action | Time |
|----------|--------|------|
| **Writing tests** | Fail loudly, no silent catches | Always |
| **New feature** | Full implementation, no NotImplementedException | Always |
| **Pool code change** | Run OnDemand tests | 3-5 min |
| **Significant task** | Update instructions | 5-10 min |
| **Any code change** | Create PR, never direct commit | Always |

---

**💡 Remember:** These rules exist because they prevent production bugs. Follow them religiously.
