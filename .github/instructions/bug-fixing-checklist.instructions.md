---
applyTo: "**/*.cs,**/*.md"
---

# Bug Fixing Checklist

> **Complete checklist for fixing bugs effectively and comprehensively**

## 🐛 Bug Fix Process (6 Steps)

### Step 1: Understand the Bug (Root Cause Analysis)

**Actions:**
1. ✅ **Read the bug report** - Understand what the user expected vs what happened
2. ✅ **Reproduce the issue** - Create minimal reproduction case
3. ✅ **Find the code** - Locate ALL files involved (tool → core → helpers)
4. ✅ **Trace the flow** - Follow parameter passing from entry point to implementation
5. ✅ **Identify root cause** - What's missing, wrong, or ignored?

**Example (refresh + loadDestination bug):**
- User expected: `refresh` with `loadDestination='worksheet'` should load data
- Actual: Parameter was silently ignored
- Root cause: `RefreshPowerQueryAsync()` didn't accept `loadDestination` parameter
- Traced: `ExcelPowerQueryTool.ExcelPowerQuery()` → `RefreshPowerQueryAsync()` → `commands.RefreshAsync()`

**Common Patterns:**
- **Parameter ignored** - Method signature doesn't include parameter
- **Wrong layer** - Logic in wrong place (MCP vs Core vs CLI)
- **Missing validation** - No error when invalid input provided
- **Silent failures** - Exceptions caught without reporting
- **Incomplete implementation** - Feature partially implemented

---

### Step 2: Fix the Code (Minimal Changes)

**Actions:**
1. ✅ **Update method signatures** - Add missing parameters to ALL layers
2. ✅ **Wire parameters through** - Ensure parameters flow from tool → core
3. ✅ **Implement logic** - Add the missing/broken functionality
4. ✅ **Handle errors** - Proper error messages for invalid states
5. ✅ **Test locally** - Verify fix works with minimal test case

**Principles:**
- ✅ Make **smallest possible changes** to fix the issue
- ✅ Maintain **backwards compatibility** (optional parameters, default values)
- ✅ Follow **existing patterns** in the codebase
- ✅ Add logic at **correct layer** (Core for business logic, MCP for tool wiring)

**Example (refresh fix):**
```csharp
// BEFORE: Parameter not accepted
private static async Task<string> RefreshPowerQueryAsync(
    PowerQueryCommands commands, 
    string excelPath, 
    string? queryName, 
    string? batchId)  // ❌ Missing loadDestination

// AFTER: Parameter accepted and used
private static async Task<string> RefreshPowerQueryAsync(
    PowerQueryCommands commands, 
    string excelPath, 
    string? queryName, 
    string? loadDestination,  // ✅ Added
    string? targetSheet,      // ✅ Added
    string? batchId)
{
    // ✅ Added logic to apply load config if specified
    if (!string.IsNullOrEmpty(loadDestination)) {
        // Apply load configuration before refresh
    }
}
```

---

### Step 3: Add Comprehensive Tests (MANDATORY)

**Actions:**
1. ✅ **Regression test** - Test the exact scenario from bug report
2. ✅ **Edge cases** - Test all variations of the fix
3. ✅ **Backwards compatibility** - Test existing behavior still works
4. ✅ **Error cases** - Test invalid inputs, failures
5. ✅ **Integration tests** - Test end-to-end (MCP Server layer)

**Test Coverage Requirements:**

**Core Layer Tests (Business Logic):**
```csharp
// File: tests/ExcelMcp.Core.Tests/Integration/Commands/Feature/FeatureCommandsTests.NewFeature.cs

[Fact]
public async Task BugFix_ExactScenarioFromReport_WorksCorrectly()
{
    // Arrange - Reproduce exact bug scenario
    // Act - Execute the fix
    // Assert - Verify it works as expected
}

[Fact]
public async Task BugFix_EdgeCase1_HandledCorrectly() { }

[Fact]
public async Task BugFix_EdgeCase2_HandledCorrectly() { }

[Fact]
public async Task BugFix_BackwardsCompatibility_ExistingCodeStillWorks() { }
```

**MCP Server Layer Tests (End-to-End):**
```csharp
// File: tests/ExcelMcp.McpServer.Tests/Integration/Tools/FeatureToolTests.cs

[Fact]
public async Task BugFix_EndToEnd_RegressionTest()
{
    // Test the exact JSON/parameter flow from bug report
}
```

**Minimum Test Count:**
- ✅ At least **3-5 Core layer tests** (regression + edge cases + backwards compat)
- ✅ At least **2-3 MCP Server tests** (end-to-end validation)
- ✅ **Total: 5-8 new tests minimum** for a bug fix

**Example (refresh bug - 13 tests):**
- 7 Core layer tests (all loadDestination values, backwards compat, custom sheet)
- 6 MCP Server tests (end-to-end with JSON serialization)

---

### Step 4: Update Documentation (3 Files Minimum)

**Actions:**
1. ✅ **Parameter descriptions** - Update tool/method XML comments
2. ✅ **User documentation** - Update COMMANDS.md or README.md
3. ✅ **Workflow hints** - Update SuggestedNextActions in result objects
4. ✅ **LLM prompts** - Update prompt files to teach LLMs about the fix
5. ✅ **Error messages** - Add helpful hints in error cases

**Files to Update:**

**1. Tool/Method Documentation:**
```csharp
// src/ExcelMcp.McpServer/Tools/FeatureTool.cs

/// <summary>
/// UPDATED: Now supports newParameter for enhanced functionality
/// </summary>
[Description("Updated description mentioning new capability")]
string? newParameter = null
```

**2. User-Facing Documentation:**
```markdown
# docs/COMMANDS.md

**command-name** - Description

Now supports --new-flag to enable new behavior:
```bash
excelcli command-name --new-flag value
```

For MCP Server users:
```javascript
tool(action: "action", newParameter: "value")
```
```

**3. Workflow Hints (SuggestedNextActions):**
```csharp
// Update result messages to reflect new capability
result.SuggestedNextActions = newFeatureUsed
    ? ["Try the new feature next", "Verify results"]
    : ["Consider using newParameter for enhanced workflow"];
```

**4. LLM Prompts:**
```csharp
// src/ExcelMcp.McpServer/Prompts/FeaturePrompts.cs

NEW FEATURE (BUGFIX):
- Now supports newParameter to enable XYZ
- Example: tool(action: "action", newParameter: "value")
- Use when: user wants to do XYZ in one call
```

**Minimum Documentation Updates:**
- ✅ At least **3 files** (tool docs, user docs, prompts)
- ✅ **SuggestedNextActions** enhanced with new capability
- ✅ **Error messages** include helpful hints

---

### Step 5: Verify Quality (Build + Tests + Checklist)

**Build Verification:**
```bash
# Clean build
dotnet build -c Release

# Verify 0 warnings, 0 errors
```

**Test Verification:**
```bash
# Run all unit tests (must pass)
dotnet test --filter "Category=Unit&RunType!=OnDemand"

# Verify all new tests are included and pass
dotnet test --filter "FullyQualifiedName~NewFeatureTests"
```

**Quality Checklist:**
- ✅ **Build passes** with 0 warnings, 0 errors
- ✅ **All unit tests pass** (141+ tests)
- ✅ **New tests created** (5-8 minimum)
- ✅ **Documentation updated** (3+ files)
- ✅ **Backwards compatible** (existing code works)
- ✅ **No TODO/FIXME markers** left in code
- ✅ **Proper error handling** with helpful messages
- ✅ **Code follows existing patterns** (batch API, async/await, etc.)

---

### Step 6: Create Summary Documentation

**Actions:**
1. ✅ **Bug fix summary** - Explain what was broken and how it's fixed
2. ✅ **Test coverage summary** - Document all new tests
3. ✅ **Documentation summary** - List all doc changes
4. ✅ **User impact** - Explain workflow improvements

**Create 3 Summary Files:**

**1. BUG-FIX-[FEATURE].md:**
```markdown
# Bug Fix: [Feature Name]

## Problem Report
User reported: [exact issue]

## Root Cause
[Technical explanation]

## Solution
[What was changed]

## Behavior Changes
Before: [old behavior]
After: [new behavior]

## Backwards Compatibility
✅ Fully backwards compatible
```

**2. TESTS-[FEATURE].md:**
```markdown
# Test Coverage for [Feature] Bug Fix

## Summary
Added X tests covering Y scenarios

## Test Files Created
1. Core layer: [file path] (X tests)
2. MCP Server: [file path] (Y tests)

## Test Scenarios
- Scenario 1: [description]
- Scenario 2: [description]
```

**3. DOCS-[FEATURE].md:**
```markdown
# Documentation Updates for [Feature]

## Files Updated
1. Tool documentation
2. User documentation  
3. Prompts

## User-Facing Changes
[Workflow improvements]
```

---

## 🚨 Common Mistakes to Avoid

### ❌ Mistake 1: Fixing Code Without Tests
**Problem:** Bug might come back, no regression protection  
**Solution:** Always add tests BEFORE marking bug as fixed

### ❌ Mistake 2: Updating Code But Not Docs
**Problem:** Users don't know about the fix or new capability  
**Solution:** Update docs in same PR as code fix

### ❌ Mistake 3: Not Updating Workflow Hints
**Problem:** Users get stale suggestions that don't mention new features  
**Solution:** Update SuggestedNextActions and WorkflowHint messages

### ❌ Mistake 4: Testing Only the Happy Path
**Problem:** Edge cases and errors still broken  
**Solution:** Test all variations, error cases, backwards compatibility

### ❌ Mistake 5: Breaking Backwards Compatibility
**Problem:** Existing user code breaks  
**Solution:** Make parameters optional, use defaults, maintain existing behavior

### ❌ Mistake 6: Ignoring Parameter Flow
**Problem:** Parameter accepted but not used (original bug!)  
**Solution:** Trace parameter from tool → implementation, ensure it's wired correctly

### ❌ Mistake 7: Fixing Symptoms, Not Root Cause
**Problem:** Bug appears in different form later  
**Solution:** Understand WHY it broke, not just WHAT is broken

### ❌ Mistake 8: No Summary Documentation
**Problem:** Hard to understand what changed and why  
**Solution:** Create summary docs explaining fix, tests, and impact

---

## ✅ Bug Fix Quality Checklist

**Before Marking Bug as Fixed:**

### Code Changes
- [ ] Root cause identified and documented
- [ ] Minimal code changes (surgical fix)
- [ ] Parameters wired through all layers
- [ ] Error handling with helpful messages
- [ ] Backwards compatible (optional params, defaults)
- [ ] Follows existing code patterns
- [ ] No TODO/FIXME markers left

### Tests
- [ ] Regression test for exact bug scenario
- [ ] Edge case tests (3-5 variations)
- [ ] Backwards compatibility test
- [ ] Error case tests
- [ ] MCP Server end-to-end tests
- [ ] Minimum 5-8 new tests total
- [ ] All tests pass (including existing tests)

### Documentation
- [ ] Tool/method XML comments updated
- [ ] Parameter descriptions updated
- [ ] User documentation updated (COMMANDS.md)
- [ ] SuggestedNextActions enhanced
- [ ] WorkflowHint messages updated
- [ ] Error messages include helpful hints
- [ ] LLM prompts updated (if applicable)
- [ ] Minimum 3 files updated

### Quality
- [ ] Build passes (0 warnings, 0 errors)
- [ ] All unit tests pass (141+ tests)
- [ ] New tests execute successfully
- [ ] No regressions in existing functionality
- [ ] Summary documentation created (3 files)

### PR Readiness
- [ ] Branch created (not main)
- [ ] Commit messages descriptive
- [ ] PR description includes bug report link
- [ ] Summary docs included in PR
- [ ] Ready for review

---

## 📊 Bug Fix Metrics

**Good Bug Fix:**
- ✅ 1 bug report → 5-8 new tests → 3+ doc files updated
- ✅ Backwards compatible (0 breaking changes)
- ✅ Build passing, all tests green
- ✅ User workflow improved (fewer steps)

**Example: Refresh + LoadDestination Bug:**
- 📝 1 bug report (user issue)
- 🐛 1 root cause (parameter ignored)
- 💻 2 files changed (ExcelPowerQueryTool.cs, prompts)
- ✅ 13 tests added (7 Core + 6 MCP Server)
- 📚 5 files documented (tool, COMMANDS.md, prompts, 3 summaries)
- 🎯 Result: 2 operations → 1 operation (50% workflow improvement)

---

## 🎓 Lessons Learned (Update This Section)

**Key Insights from Recent Bug Fixes:**

1. **Always trace parameters end-to-end** - Parameter must flow from tool → implementation
2. **Update workflow hints** - Users rely on SuggestedNextActions for guidance
3. **Test all layers** - Core logic AND MCP Server end-to-end
4. **Document for LLMs** - Update prompts to teach AI assistants about new capabilities
5. **Create summary docs** - Makes PR review and future reference easier

---

## 📖 Related Documentation

- [Critical Rules](critical-rules.instructions.md) - Mandatory development rules
- [Testing Strategy](testing-strategy.instructions.md) - Test architecture and patterns
- [Development Workflow](development-workflow.instructions.md) - PR process and CI/CD
- [MCP Server Guide](mcp-server-guide.instructions.md) - MCP tool patterns
