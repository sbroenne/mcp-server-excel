---
applyTo: "**"
---

# CRITICAL RULES - MUST FOLLOW

> **⚠️ NON-NEGOTIABLE rules for all ExcelMcp development**

## Rule 0: NEVER Commit Without Running Tests (CRITICAL)

**NEVER commit, push, or create PRs without first running tests for the code you changed.**

**Why Critical:** Prevents breaking changes from reaching main, wastes team time debugging failures, violates CI/CD principles.

**Enforcement:**
- ALWAYS run relevant tests BEFORE committing
- Use `--filter "Feature=<feature-name>&RunType!=OnDemand"` for feature-specific tests
- Build must pass (0 warnings, 0 errors)
- Pre-commit hooks must pass (COM leaks, success flag, coverage)
- Document test results in commit message

**Process:**
1. Make code changes
2. Build: `dotnet build` (must succeed with 0 warnings)
3. Run tests: `dotnet test --filter "Feature=<feature>&RunType!=OnDemand"`
4. Verify all tests pass
5. Run pre-commit checks
6. THEN commit

## Rule 21: Never Commit Automatically (CRITICAL)

**NEVER commit or push code automatically. All commits, pushes, and merges must require explicit user approval.**

**Why Critical:** Prevents accidental changes, enforces review, and ensures user control over all repository modifications.

**Enforcement:**
- All automated tools, scripts, and agents must prompt for user approval before any commit, push, or merge.
- No background or silent commits allowed.
- Document this rule in all agent and automation instructions.


## Quick Reference (Grouped by Context)

**Note:** Rules below are grouped by when you need them, not by number. Detailed rules follow in numeric order (1-21).

**Every Edit:**
| Rule | Action | Why Critical |
|------|--------|--------------|
| 0. Test before commit | ALWAYS run tests before committing | Prevents breaking changes |
| 1. Success flag | NEVER `Success=true` with `ErrorMessage` | Confuses LLMs, causes silent failures |
| 16. Test scope | Only run tests for code you changed | Saves 10+ minutes per test run |
| 8. TODO markers | Must resolve before commit | Pre-commit hook blocks |

**When Writing Code:**
| Rule | Action | Why Critical |
|------|--------|--------------|
| 7. COM API | Use Excel COM first, validate docs | Prevents wrong dependencies |
| 9. GitHub search | Search OTHER repos for VBA/COM examples FIRST | Learn from working code |
| 2. NotImplementedException | Never use, full implementation only | No placeholders allowed |
| 15. Enum mappings | All enum values mapped in ToActionString() | Runtime errors otherwise |
| 17. MCP error checks | Always return JSON, check result.Success | MCP protocol requirement |

**When Writing Tests:**
| Rule | Action | Why Critical |
|------|--------|--------------|
| 2. Tests | Fail loudly, never silent | Silent failures waste hours |
| 15. No Save | Remove unless testing persistence | Makes tests 50% faster |
| 11. Test debugging | Run tests one by one | Isolates actual failure |
| 13. Test compliance | Pass checklist before PR submission | Prevents test pollution |

**Before Commit:**
| Rule | Action | Time |
|------|--------|------|
| 0. Test before commit | ALWAYS run tests before committing | 3-5 min |
| 4. Session code | See testing-strategy for test commands | 3-5 min |
| 6. COM leaks | Pre-commit hook auto-checks | 1 min |
| 7. PRs | Always use PRs, never direct commit | Always |

**During PR Process:**
| Rule | Action | Time |
|------|--------|------|
| 20. PR review comments | Check/fix automated comments immediately | 5-10 min |
| 14. Bug fixes | Complete 6-step process | 30-60 min |
| 19. Tool descriptions | Verify [Description] matches behavior | Per change |

**Rare/Specialized:**
| Rule | Action | When |
|------|--------|------|
| 12. No test refs | Production NEVER references tests | Architecture review |
| 5. Instructions | Update after significant work | After major features |

---

## Rule 1: Success Flag Must Match Reality (CRITICAL)

**NEVER set `Success = true` when `ErrorMessage` is set. This is EXTREMELY serious!**

```csharp
// ❌ CRITICAL BUG: Confuses LLMs and users
result.Success = true;
result.ErrorMessage = "Query imported but failed to load...";

// ✅ CORRECT: Success only when NO errors
if (!loadResult.Success) {
    result.Success = false;  // MUST be false!
    result.ErrorMessage = $"Failed: {loadResult.ErrorMessage}";
}
```

**Invariant:** `Success == true` ⟹ `ErrorMessage == null || ErrorMessage == ""`

**Why Critical:** LLMs see Success=true and assume operation worked, causing workflow failures and silent data corruption.

**Common Bug Pattern (43 violations found 2025-01-28):**
```csharp
// ❌ WRONG: Optimistic Success setting without catch block correction
var result = new OperationResult();
result.Success = true;  // Set optimistically

try {
    // ... do work ...
    return result;
} catch (Exception ex) {
    // ❌ BUG: Forgot to set Success = false!
    result.ErrorMessage = $"Error: {ex.Message}";
    return result;  // Returns Success=true with ErrorMessage! 
}

// ✅ CORRECT: Set Success in try block, always false in catch
var result = new OperationResult();

try {
    // ... do work ...
    result.Success = true;  // Only set true on actual success
    return result;
} catch (Exception ex) {
    result.Success = false;  // ✅ Always false in catch!
    result.ErrorMessage = $"Error: {ex.Message}";
    return result;
}
```

**Enforcement:**
- Pre-commit hook runs `check-success-flag.ps1` to detect violations
- Regression tests verify this invariant (PowerQuerySuccessErrorRegressionTests)
- Code review MUST check every `Success = ` assignment
- Search pattern: `Success.*true.*ErrorMessage`

**Examples of bugs found:**
- 43 violations across Connection, PowerQuery, DataModel, VBA, Range, Table commands
- All followed pattern: `Success = true` at start, `ErrorMessage` set in catch without `Success = false`

---

## Rule 2: No NotImplementedException

**Every feature must be fully implemented with real Excel COM operations and passing tests. No placeholders.**

---

## Rule 3: Session Cleanup Tests

**When modifying session/batch code, run OnDemand tests (see testing-strategy.instructions.md). Must pass before commit.**

---

## Rule 4: Update Instructions

**After significant work, update `.github/copilot-instructions.md` with lessons learned, architecture changes, and testing insights.**

---

## Rule 5: COM Object Leak Detection

**Before commit: `& "scripts\check-com-leaks.ps1"` must report 0 leaks.**

All `dynamic` COM objects must be released in `finally` blocks using `ComUtilities.Release(ref obj!)`.

Exception: Session management files (ExcelBatch.cs, ExcelSession.cs).

---

## Rule 6: All Changes Via Pull Requests

**Never commit to `main`. Create feature branch → PR → CI/CD + review → merge.**

**Enforcement:** Pre-commit hook blocks commits to main. If you're on main:
```bash
git stash                                    # Save changes
git checkout -b feature/your-feature-name    # Create feature branch
git stash pop                                # Restore changes
git add <files>                              # Stage changes
git commit -m "your message"                 # Commit to feature branch
```

**Why Critical:** Direct commits to main bypass CI/CD, skip code review, and violate branch protection.

---

## Rule 7: COM API First

**Use Excel COM API for everything it supports. Only use external libraries (TOM) for features Excel COM doesn't provide.**

Validate against [Microsoft docs](https://learn.microsoft.com/office/vba/api/overview/excel) before adding dependencies.

---

## Rule 8: No TODO/FIXME Markers

**Code must be complete before commit. No TODO, FIXME, HACK, or XXX markers in source code.**

Delete commented-out code (use git history). Exception: Documentation files only.

---

## Rule 9: Search External GitHub Repositories for Working Examples First

**BEFORE creating new Excel COM Interop code or troubleshooting COM issues:**

- **ALWAYS** search OTHER open source GitHub repositories for working examples
- **NEVER** search your own repository - only search external projects
- **NetOffice is THE BEST source for ALL COM Interop work**: https://github.com/NetOfficeFw/NetOffice
  - Strongly-typed C# wrappers for ALL Office COM APIs (Excel, Word, PowerPoint, Outlook, etc.)
  - Search for ANY Excel COM operation: ranges, worksheets, PivotTables, Power Query, charts, VBA, connections, formatting, etc.
  - Study their patterns for dynamic interop conversion and proper COM object handling
  - NetOffice source code is essentially a comprehensive reference for every Excel COM API
- Look for repositories with Excel automation, VBA code, or Office interop projects
- Search for the specific COM object/method you need (e.g., "PivotTable CreatePivotTable VBA", "QueryTable Refresh VBA", "Range.Value2 NetOffice")
- Study proven patterns from other projects before writing new code
- Avoid reinventing solutions - learn from working implementations in the wild

**Why:** Excel COM is quirky. Real-world VBA examples from other projects prevent common pitfalls (1-based indexing, object cleanup, async issues, variant types, etc.)

---

## Rule 10: Debug Tests One by One

**When debugging test failures, ALWAYS run tests individually - never run all tests at once.**

**Process:**
1. List all test methods in the file
2. Run each test individually using `--filter "FullyQualifiedName=Namespace.Class.Method"`
3. Identify exact failure for each test before moving to next
4. Fix issues one test at a time

**Why:** Running all tests together masks which specific test fails and why. Individual execution provides clear, isolated diagnostics.

---

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

---

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
6. ✅ **PR Description** - Comprehensive summary (bug report, fix, tests, docs updated)

**Process:** Follow [bug-fixing-checklist.instructions.md](bug-fixing-checklist.instructions.md) for complete 6-step process.

**Why:** Incomplete bug fixes lead to regressions, confusion, and wasted time. Comprehensive fixes prevent future issues.

**Example:** Refresh + loadDestination bug = 1 code file + 13 tests + 5 doc files + detailed PR description = complete fix.

---

## Rule 14: No Save Unless Testing Persistence

**Code must NOT call `batch.Save()` unless explicitly testing persistence.**

**Quick Rules:**
- ❌ FORBIDDEN: Tests only verifying operation success or in-memory state
- ✅ REQUIRED: Round-trip tests verifying data persists after workbook close/reopen
- ⚡ REASON: Save is slow (~2-5s). Removing unnecessary saves makes tests 50%+ faster

**See:** [testing-strategy.instructions.md](testing-strategy.instructions.md) for complete Save patterns, when to use, and detailed examples.

---

## Quick Reference

| Rule | Action | Time |
|------|--------|------|
| 0. Test before commit | ALWAYS run tests before committing | 3-5 min |
| 1. Tests | Fail loudly, never silent | Always |
| 2. NotImplementedException | Never use, full implementation only | Always |
| 3. Session code | See testing-strategy for test commands | 3-5 min |
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
| 14. No Save | See testing-strategy for complete patterns | Per test |
| 15. Enum mappings | All enum values mapped in ToActionString() | Always |
| 16. Test scope | Only run tests for code you changed | Per change |
| 17. MCP error checks | Check result.Success before JsonSerializer.Serialize | Every method |
| 18. Tool descriptions | Verify [Description] matches tool behavior | Per tool change |
| 19. PR review comments | Check and fix all automated review comments after creating PR | 5-10 min |



---

## Rule 15: Complete Enum Mappings (CRITICAL)

**Every enum value MUST have a mapping in ToActionString(). Missing mappings cause unhandled exceptions.**

```csharp
// ❌ WRONG: Incomplete mapping
public static string ToActionString(this RangeAction action) => action switch
{
    RangeAction.GetValues => "get-values",
    RangeAction.SetValues => "set-values",
    // Missing GetUsedRange, GetCurrentRegion, etc. → ArgumentException!
    _ => throw new ArgumentException($"Unknown RangeAction: {action}")
};

// ✅ CORRECT: All enum values mapped
public static string ToActionString(this RangeAction action) => action switch
{
    RangeAction.GetValues => "get-values",
    RangeAction.SetValues => "set-values",
    RangeAction.GetUsedRange => "get-used-range",  // ✅ All values
    RangeAction.GetCurrentRegion => "get-current-region",
    // ... all other values
    _ => throw new ArgumentException($"Unknown RangeAction: {action}")
};
```

**Why Critical:** Missing mappings cause MCP Server to throw exceptions instead of returning JSON, confusing LLMs.

**Enforcement:**
- Regression tests for all enum mappings
- When adding enum value, add mapping immediately
- Code review MUST verify completeness

**Example Bug:** `GetUsedRange` missing → "An error occurred invoking 'excel_range'" (not JSON!)

---

## Rule 16: Test Only What You Changed (CRITICAL - PERFORMANCE)

**ALWAYS run tests ONLY for the specific code you modified. Integration tests take a very long time.**

**Wrong:**
```bash
# ❌ NEVER: Runs ALL integration tests (10+ minutes)
dotnet test --filter "Category=Integration&RunType!=OnDemand"
```

**Correct:**
```bash
# ✅ CORRECT: Test only the feature you changed
dotnet test --filter "Feature=PowerQuery&RunType!=OnDemand"  # PowerQuery changes only
dotnet test --filter "Feature=Connection&RunType!=OnDemand"  # Connection changes only
dotnet test --filter "Feature=Sheet&RunType!=OnDemand"       # Sheet changes only
```

**Why Critical:** Integration tests require Excel COM automation and are SLOW. Running all tests wastes time and resources.

**Enforcement:**
- Only run tests for files you modified
- Use Feature trait to target specific test groups
- Full test suite runs in CI/CD pipeline only

---

## Rule 17: MCP Tools Must Return JSON Responses (CORRECTED)

**Every MCP tool method that calls Core Commands MUST return JSON responses, not throw exceptions for business errors.**

```csharp
// ❌ WRONG: Throws exception for business logic errors
private static async Task<string> SomeAction(...)
{
    var result = await commands.SomeAsync(batch, param);
    
    if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
    {
        throw new ModelContextProtocol.McpException($"action failed: {result.ErrorMessage}");  // ❌ Wrong!
    }
    
    return JsonSerializer.Serialize(result, JsonOptions);
}

// ✅ CORRECT: Always return JSON - let result.Success indicate errors
private static async Task<string> SomeAction(...)
{
    var result = await commands.SomeAsync(batch, param);
    
    // Always return JSON (success or failure) - MCP clients handle the success flag
    return JsonSerializer.Serialize(result, JsonOptions);
}
```

**When to Throw McpException:**
- ✅ **Parameter validation** - missing required params, invalid formats
- ✅ **Pre-conditions** - file not found, batch not found, invalid state
- ❌ **NOT for business logic errors** - table not found, query failed, etc.

**Why:**
- ✅ MCP clients expect JSON responses with `success: false` for business errors
- ✅ HTTP 200 + JSON error = client can parse and handle gracefully
- ❌ HTTP 500 + exception = harder for clients to handle programmatically
- ✅ Core Commands return result objects with `Success` flag - serialize them!

**Example - Business Error (return JSON):**
```csharp
// Core returns: { Success = false, ErrorMessage = "Table 'Sales' not found" }
// MCP Tool: Return this as-is
return JsonSerializer.Serialize(result, JsonOptions);
// Client gets: {"success": false, "errorMessage": "Table 'Sales' not found"}
```

**Example - Validation Error (throw exception):**
```csharp
// Missing required parameter
if (string.IsNullOrWhiteSpace(tableName))
{
    throw new ModelContextProtocol.McpException("tableName is required for create-from-table action");
}
```

**Historical Note:** This rule was corrected on 2025-01-03 after discovering that tests expected JSON responses, not exceptions. The previous pattern (throwing McpException for business errors) was incorrect and caused MCP clients to receive unhandled errors instead of parseable JSON.

---

## Rule 18: Tool Descriptions Must Match Behavior (CRITICAL)

**Tool `[Description]` attributes are part of the MCP schema sent to LLMs. They must be accurate and current.**

**What to verify when changing a tool:**

1. **Purpose and Use Cases Clear**:
   ```csharp
   // ❌ WRONG: Vague description
   [Description("Manage worksheets")]
   
   // ✅ CORRECT: Clear purpose and use cases
   [Description("Manage Excel worksheet lifecycle: create, rename, copy, delete sheets")]
   ```

2. **Non-Enum Parameter Values Documented**:
   ```csharp
   // ❌ WRONG: Parameter values not explained
   [Description("Import Power Query with loadDestination parameter")]
   
   // ✅ CORRECT: Non-enum parameter values explained
   [Description(@"Import Power Query.
   
   LOAD DESTINATIONS:
   - 'worksheet': Load to worksheet (DEFAULT)
   - 'data-model': Load to Power Pivot
   - 'both': Load to BOTH
   - 'connection-only': Don't load data")]
   ```

3. **Server-Specific Behavior Documented**:
   ```csharp
   // ❌ WRONG: Behavior changed but description outdated
   [Description("Default: loadDestination='connection-only'")]  // Wrong!
   
   // ✅ CORRECT: Description reflects actual default
   [Description("Default: loadDestination='worksheet'")]
   ```

**What NOT to include:**
- ❌ **Enum action lists** - MCP SDK auto-generates these in schema (LLMs see them via dropdown)
- ❌ **Parameter types** - Schema provides this
- ❌ **Required/optional flags** - Schema provides this

**Why Critical:** LLMs use tool descriptions for server-specific guidance. Inaccurate descriptions cause:
- Wrong default parameter values
- Incorrect workflow assumptions
- Confused users when behavior doesn't match docs

**When to Update:**
- Changing default values or server behavior
- Adding/changing non-enum parameter values (loadDestination, formatCode, etc.)
- Changing which tools to use for related operations
- Adding performance guidance (batch mode)

**See:** [mcp-server-guide.instructions.md](mcp-server-guide.instructions.md) for complete Tool Description checklist.

---

## Rule 19: Check PR Review Comments After Creating PR (CRITICAL)

**After creating a PR, ALWAYS check for automated review comments from Copilot and GitHub Advanced Security.**

```bash
# Retrieve inline code review comments using GitHub CLI
gh api repos/sbroenne/mcp-server-excel/pulls/PULL_NUMBER/comments --paginate

# Or use the mcp_github tool if available
mcp_github_github_pull_request_read(method="get_review_comments", owner="sbroenne", repo="mcp-server-excel", pullNumber=PULL_NUMBER)
```

**Common automated reviewers:**
- **Copilot** (code quality, performance, style)
- **github-advanced-security** (security scanning, code analysis)

**Common issues to fix:**
- Improper `/// <inheritdoc/>` on constructors/test methods that don't override
- `.AsSpan().ToString()` inefficiency - use `[..n]` range operator instead
- Nullable type access without null checks
- `foreach` → `.Select()` for functional style
- Nested if statements that can be combined
- Generic catch clauses - use specific exceptions or add justification
- Path.Combine security warnings - suppress with justification for test code

**Fix all automated review comments before requesting human review.**

**Why Critical:** Automated reviewers catch common code quality issues early. Fixing them promptly:
- Improves code quality and maintainability
- Reduces human reviewer workload
- Speeds up PR approval process
- Prevents accumulation of technical debt

**Process:**
1. Create PR
2. Immediately check for review comments (within 1-2 minutes)
3. Fix all automated issues in a single follow-up commit
4. Push fixes to PR branch
5. Request human review only after all automated issues resolved

**Example:** PR #139 had 17 automated review comments - all fixed in one commit before human review.

---
---
