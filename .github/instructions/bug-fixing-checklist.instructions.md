---
applyTo: "**/*.cs,**/*.md"
---

# Bug Fixing Checklist

> **6-step process for comprehensive bug fixes**

## Process

1. **Root Cause** - Trace flow from entry point to bug, identify what's missing/wrong/ignored
2. **Fix Code** - Minimal changes at correct layer, maintain backwards compatibility
3. **Add Tests** - Minimum 5-8 tests: regression + edge cases + backwards compat + MCP end-to-end
4. **Update Docs** - Minimum 3 files: tool/method docs, user docs, SuggestedNextActions, LLM prompts
5. **Verify Quality** - Build passes (0 warnings), all tests pass, no TODOs left
6. **PR Description** - Bug summary, root cause, fix explanation, test coverage, docs updated

## Test Coverage Requirements

**Core Layer** (3-5 tests):
- Exact bug scenario (regression)
- Edge case variations
- Backwards compatibility validation

**MCP Server Layer** (2-3 tests):
- End-to-end JSON serialization flow
- Parameter passing verification

**Total: 5-8 new tests minimum**

## Documentation Requirements

**Required files (minimum 3)**:
1. Tool/method XML documentation (`/// <summary>`, `/// <param>`)
2. User-facing docs (README or component docs)
3. LLM prompts in `src/ExcelMcp.McpServer/Prompts/Content/`

**Update workflow hints**:
- `SuggestedNextActions` - reflect new capability
- Error messages - include helpful hints
- `WorkflowHint` - guide next steps

## Quality Checklist

**Before marking bug as fixed**:
- [ ] Root cause documented
- [ ] Minimal code changes (surgical fix)
- [ ] Parameters wired through all layers
- [ ] 5-8 new tests added (Core + MCP)
- [ ] 3+ doc files updated
- [ ] Build passes (0 warnings, 0 errors)
- [ ] All tests pass (including existing)
- [ ] No TODO/FIXME markers
- [ ] Backwards compatible
- [ ] PR description comprehensive

## Common Mistakes

| Mistake | Fix |
|---------|-----|
| Code without tests | Add tests BEFORE marking fixed |
| Code without docs | Update docs in same PR |
| Happy path only | Test edge cases, errors, backwards compat |
| Breaking changes | Make params optional, use defaults |
| Parameter ignored | Trace from tool → implementation |
| Symptoms fixed, not root cause | Understand WHY it broke |
| Incomplete PR | Document bug, fix, tests, docs updated |

## PR Description Template

```markdown
## Bug Fix: [Feature Name]

### Problem
User reported: [exact issue]
Issue: #[number]

### Root Cause
[Technical explanation]

### Solution
**Files Changed:**
- path/to/file.cs - [what changed]

**Behavior:**
- Before: [old behavior]
- After: [new behavior]

### Test Coverage (X tests, Y scenarios)
**Core:** test1, test2, test3
**MCP Server:** test4, test5

**Test Files:**
- tests/.../FeatureCommandsTests.NewFeature.cs
- tests/.../FeatureToolTests.cs

### Documentation Updated
1. MCP prompts - [what updated]
2. Tool/method docs - [what updated]
3. User docs - [what updated]

### Backwards Compatibility
✅ Fully backwards compatible - [how]

### User Impact
[Workflow improvements]
```

**No separate summary files** - PR description is canonical record (searchable via GitHub)
