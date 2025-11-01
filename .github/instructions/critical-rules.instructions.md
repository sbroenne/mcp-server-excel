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



## Rule 3: Session Cleanup Tests

When modifying session/batch code (`ExcelSession.cs`, `ExcelBatch.cs`, `ExcelHelper.cs`), run: `dotnet test --filter "RunType=OnDemand"`
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



## Rule 9: Search External GitHub Repositories for Working Examples First

**BEFORE** creating new Excel COM Interop code or troubleshooting COM issues:
- **ALWAYS** search OTHER open source GitHub repositories for working examples
- **NEVER** search your own repository - only search external projects
- Look for repositories with Excel automation, VBA code, or Office interop projects
- Search for the specific COM object/method you need (e.g., "PivotTable CreatePivotTable VBA", "QueryTable Refresh VBA", "ModelMeasures.Add VBA")
- Study proven patterns from other projects before writing new code
- Avoid reinventing solutions - learn from working implementations in the wild

**Why:** Excel COM is quirky. Real-world VBA examples from other projects prevent common pitfalls (1-based indexing, object cleanup, async issues, variant types, etc.)



## Rule 10: Debug Tests One by One

When debugging test failures, **ALWAYS run tests individually** - never run all tests at once.

**Process:**
1. List all test methods in the file
2. Run each test individually using `--filter "FullyQualifiedName=Namespace.Class.Method"`
3. Identify exact failure for each test before moving to next
4. Fix issues one test at a time

**Why:** Running all tests together masks which specific test fails and why. Individual execution provides clear, isolated diagnostics.



## Rule 11: Production Code NEVER References Tests

**Production code (Core, CLI, MCP Server) must NEVER reference test projects or test helpers.**

**Violations:**
- ❌ `<InternalsVisibleTo Include="*.Tests" />` in production `.csproj`
- ❌ `using Sbroenne.ExcelMcp.*.Tests` in production code
- ❌ Production code calling test helper methods
- ❌ Production business logic in helper classes that tests use

**Correct Architecture:**
- ✅ **COM utilities** → `ComInterop/ComUtilities.cs` (low-level COM helpers like SafeGetString, ForEach iterators)
- ✅ **Business logic** → Private methods inside production Commands classes
- ✅ **Test helpers** → Call production commands, never duplicate logic
- ✅ `InternalsVisibleTo` only for production-to-production (e.g., Core → MCP Server)

**Why:** Tests depend on production code, not the reverse. Production code with test dependencies is broken architecture.



## Rule 12: Test Class Compliance Checklist

**Every new test class MUST pass the compliance checklist before PR submission.**

**Verify:**
- ✅ Uses `IClassFixture<TempDirectoryFixture>` (NOT manual IDisposable)
- ✅ Each test creates unique file via `CoreTestHelper.CreateUniqueTestFileAsync()`
- ✅ NEVER shares test files between tests
- ✅ VBA tests use `.xlsm` extension (NOT .xlsx renamed)
- ✅ Binary assertions only (NO "accept both" patterns)
- ✅ All required traits present (Category, Speed, Layer, RequiresExcel, Feature)
- ✅ Batch API pattern used correctly (no ValueTask.FromResult wrapper)
- ✅ NO duplicate helper methods (use CoreTestHelper)

**Why:** Systematic compliance prevents test pollution, file lock issues, silent failures, and maintenance nightmares. See [testing-strategy.instructions.md](testing-strategy.instructions.md) for complete checklist.

**Enforcement:** PR reviewers MUST check compliance before approval.

---

## Rule 13: Comprehensive Bug Fixes

**Every bug fix MUST include all 6 components before PR submission.**

**Required Components:**
1. ✅ **Code Fix** - Minimal surgical changes to fix root cause
2. ✅ **Tests** - Minimum 5-8 new tests (regression + edge cases + backwards compat)
3. ✅ **Documentation** - Update 3+ files (tool docs, user docs, prompts)
4. ✅ **Workflow Hints** - Update SuggestedNextActions and error messages
5. ✅ **Quality Verification** - Build passes, all tests green, 0 warnings
6. ✅ **Summary Docs** - Create BUG-FIX-*.md, TESTS-*.md, DOCS-*.md

**Process:** Follow [bug-fixing-checklist.instructions.md](bug-fixing-checklist.instructions.md) for complete 6-step process.

**Why:** Incomplete bug fixes lead to regressions, confusion, and wasted time. Comprehensive fixes prevent future issues.

**Example:** Refresh + loadDestination bug = 1 code file + 13 tests + 5 doc files + 3 summaries = complete fix.

---

## Quick Reference

| Rule | Action | Time |
|------|--------|------|
| 1. Tests | Fail loudly, never silent | Always |
| 2. NotImplementedException | Never use, full implementation only | Always |
| 3. Session code | Run `dotnet test --filter "RunType=OnDemand"` | 3-5 min |
| 4. Instructions | Update after significant work | 5-10 min |
| 5. COM leaks | Run `scripts\check-com-leaks.ps1` | 1 min |
| 6. PRs | Always use PRs, never direct commit | Always |
| 7. COM API | Use Excel COM first, validate docs | Always |
| 8. TODO markers | Must resolve before commit | 1 min |
| 9. GitHub search | Search OTHER repos for VBA/COM examples FIRST | 1-2 min |
| 10. Test debugging | Run tests one by one, never all together | Per test |
| 11. No test refs | Production NEVER references tests | Always |
| 12. Test compliance | Pass checklist before PR submission | 2-3 min |
| 13. Bug fixes | Complete 6-step process (fix, test, doc, hints, verify, summarize) | 30-60 min |
