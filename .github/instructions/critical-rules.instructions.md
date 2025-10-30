---
applyTo: "**"
---

# CRITICAL RULES - MUST FOLLOW

> **⚠️ NON-NEGOTIABLE rules for all ExcelMcp development**

## Rule 1: No Silent Test Failures

Tests must fail loudly. Never catch exceptions without re-throwing or use conditional assertions that always pass.

```csharp
// ❌ WRONG: Silent failure
catch (Exception) { _output.WriteLine("Skipping"); }

// ✅ CORRECT: Fail loudly
Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
```



## Rule 2: No NotImplementedException

Every feature must be fully implemented with real Excel COM operations and passing tests. No placeholders.



## Rule 3: Pool Cleanup Tests

When modifying pool code (`ExcelInstancePool.cs`, `ExcelHelper.cs`), run: `dotnet test --filter "RunType=OnDemand"`
All tests must pass before commit. Requires Excel installed, takes 3-5 minutes.



## Rule 4: Update Instructions

After significant work, update `.github/copilot-instructions.md` with lessons learned, architecture changes, and testing insights.



## Rule 5: COM Object Leak Detection

Before commit: `& "scripts\check-com-leaks.ps1"` must report 0 leaks.
All `dynamic` COM objects must be released in `finally` blocks using `ComUtilities.Release(ref obj!)`.
Exception: Session management files (ExcelBatch.cs, ExcelSession.cs).



## Rule 6: All Changes Via Pull Requests

Never commit to `main`. Create feature branch → PR → CI/CD + review → merge.



## Rule 7: COM API First

Use Excel COM API for everything it supports. Only use external libraries (TOM) for features Excel COM doesn't provide.
Validate against [Microsoft docs](https://learn.microsoft.com/office/vba/api/overview/excel) before adding dependencies.



## Rule 8: No TODO/FIXME Markers

Code must be complete before commit. No TODO, FIXME, HACK, or XXX markers in source code.
Delete commented-out code (use git history). Exception: Documentation files only.

---

## Quick Reference

| Rule | Action | Time |
|------|--------|------|
| 1. Tests | Fail loudly, never silent | Always |
| 2. NotImplementedException | Never use, full implementation only | Always |
| 3. Pool code | Run `dotnet test --filter "RunType=OnDemand"` | 3-5 min |
| 4. Instructions | Update after significant work | 5-10 min |
| 5. COM leaks | Run `scripts\check-com-leaks.ps1` | 1 min |
| 6. PRs | Always use PRs, never direct commit | Always |
| 7. COM API | Use Excel COM first, validate docs | Always |
| 8. TODO markers | Must resolve before commit | 1 min |
