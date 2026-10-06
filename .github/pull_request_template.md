## Summary
Explain the problem, its root cause, and the intended outcome. Do not include
customer/workbook data, credentials, connection strings, or private paths.

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update
- [ ] Maintenance, tests, or CI

## Related Issues
Use `Closes #...` for an issue this change resolves, `Relates to #...` for related
work, or N/A. Do not leave placeholder issue references.

## Changeset
- [ ] Added a changeset for user-visible changes; see [.changeset/README.md](https://github.com/sbroenne/mcp-server-excel/blob/main/.changeset/README.md)
- [ ] Not applicable (internal/docs/tests/CI-only); added `skip-changelog` instead

## Changes Made
Summarize the approach, affected contracts and entry points, and any migration or
partial-state/recovery consequences. MCP Server and `excelcli` are equal entry
points; explain any adapter-only change.

## Testing Performed
Follow [AGENTS.md](https://github.com/sbroenne/mcp-server-excel/blob/main/AGENTS.md) and [tests/AGENTS.md](https://github.com/sbroenne/mcp-server-excel/blob/main/tests/AGENTS.md).
Check only completed, applicable items; explain anything not run below.

- [ ] Behavioral fix has a focused regression that failed before the fix
- [ ] Required Release build completed with zero warnings
- [ ] Affected Excel behavior passed through `scripts\Test-ExcelBehavior.ps1`; recorded its results directory
- [ ] Runtime changes in Core/ComInterop/Service/CLI/MCP or generators passed local `scripts\Test-E2E.ps1` once against final PR source
- [ ] Applicable Excel-free, contract, and source checks passed
- [ ] Assertions verify returned fields and actual Excel state, including partial state or recovery after failure
- [ ] Excel-dependent commands ran sequentially; cleanup affected only owned sessions/process identities

## Test Commands
```powershell
# Record each exact command and its result (passed, failed, or not run with reason).
# Record Test-ExcelBehavior results directories where applicable.
# -SkipBuild requires a successful Release solution build in this worktree.
```

Desktop Excel is required for COM/E2E checks; ordinary GitHub-hosted runners do
not have it. Report unavailable Excel checks as not run. Documentation/configuration-only
changes do not need synthetic runtime tests. LLM evaluations under `llm-tests/`
are on-demand only, not a normal implementation or PR gate.

## Screenshots (if applicable)
N/A unless visuals help explain the change and an interactive desktop is available.
Use only synthetic data.

## Core Commands Coverage Checklist

**Does this PR change Core contracts or generated routing?** [ ] Yes [ ] No

If yes, follow [generated contract guidance](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/agents/rules/coverage-prevention-strategy.md):

- [ ] Edited source contracts/generators, not emitted files
- [ ] Built Release and inspected generated Service, CLI options/batch JSON, and MCP schemas
- [ ] Ran `scripts\Invoke-ExcelFreeTests.ps1 -Local -Contracts`
- [ ] Verified matching names, parameters, defaults, validation, results, and timeouts across both entry points
- [ ] Updated focused tests and applicable source guidance
- [ ] Ran `scripts\check-doc-counts.ps1` (or `-SkipBuild` after a successful Release build) if advertised counts changed

## Checklist
- [ ] Self-review of code completed
- [ ] Followed AGENTS.md and matching task guides
- [ ] Updated directly affected documentation and agent-facing metadata
- [ ] Preserved errors, cancellation, and documented session recovery; no success-shaped fallback
- [ ] `Success == true` has no error message
- [ ] Applicable COM safety rules followed, including `finally` cleanup and PID/start-time process ownership
- [ ] No sensitive data in public artifacts
- [ ] No Git hooks skipped or bypassed
- [ ] Addressed and resolved review threads, or recorded a clear reason for no change

## Additional Notes
Record limitations, unavailable validation, and anything requiring careful review.
Merging or publishing requires separate authorization.
