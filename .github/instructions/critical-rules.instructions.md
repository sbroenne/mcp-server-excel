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



## Rule 9: Search Open Source Repositories for Working COM Examples First

**BEFORE** creating new Excel COM Interop code or troubleshooting COM issues:
- **ALWAYS** search other open source GitHub repositories for working examples using `github_repo` tool
- Search for the COM object/method you need (e.g., "microsoft/Excel PivotTable CreatePivotTable", "QueryTable Refresh", "ListObject Range")
- Look for repositories with Excel automation, VBA conversion, or Office interop projects
- Study proven patterns from other projects before writing new code
- Avoid reinventing solutions - learn from working implementations in the wild

**Why:** Excel COM is quirky. Real-world examples from other projects prevent common pitfalls (1-based indexing, object cleanup, async issues, variant types, etc.)



## Rule 10: Debug Tests One by One

When debugging test failures, **ALWAYS run tests individually** - never run all tests at once.

**Process:**
1. List all test methods in the file
2. Run each test individually using `--filter "FullyQualifiedName=Namespace.Class.Method"`
3. Identify exact failure for each test before moving to next
4. Fix issues one test at a time

**Why:** Running all tests together masks which specific test fails and why. Individual execution provides clear, isolated diagnostics.

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
| 9. GitHub search | Search other repos for COM examples FIRST | 1-2 min |
| 10. Test debugging | Run tests one by one, never all together | Per test |
