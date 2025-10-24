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
