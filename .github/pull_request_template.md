## Summary
Brief description of what this PR does.

## Type of Change
- [ ] 🐛 Bug fix (non-breaking change which fixes an issue)
- [ ] ✨ New feature (non-breaking change which adds functionality)
- [ ] 💥 Breaking change (fix or feature that would cause existing functionality to not work as expected)
- [ ] 📚 Documentation update
- [ ] 🔧 Maintenance (dependency updates, code cleanup, etc.)

## Related Issues
Closes #[issue number]
Relates to #[issue number]

## Changeset
- [ ] Added a changeset (`npx changeset`) describing this change for end users — see `.changeset/README.md`
- [ ] Not applicable (docs/tests/CI/dependency-only) — added `skip-changelog` label instead

## Changes Made
- Change 1
- Change 2
- Change 3

## Testing Performed
- [ ] If Core/ComInterop/Service/CLI/MCP runtime paths or their generators changed: ran `& .\scripts\Test-E2E.ps1` locally and confirmed it completed with no failures or unresolved issues
- [ ] Excel E2E not applicable because only documentation/configuration/Excel-free tooling changed
- [ ] Ran the relevant feature-specific integration tests and recorded the exact command and result below
- [ ] Tested manually with various Excel files
- [ ] Verified cleanup of this test's owned Excel sessions without terminating unrelated Excel processes (if applicable)
- [ ] Tested error conditions (missing files, invalid arguments, etc.)
- [ ] All existing commands still work
- [ ] VBA script execution tested (if applicable)
- [ ] XLSM file format validation tested (if applicable)
- [ ] VBA trust setup tested (if applicable)
- [ ] Build produces zero warnings

## Test Commands
```powershell
# Record each exact command and its result (passed, failed, or not run with reason).
# Configuration/docs-only work: record the relevant checks; Excel tests are not required.
```

## Screenshots (if applicable)
[Add screenshots showing the new functionality]

## Core Commands Coverage Checklist ⚠️

**Does this PR add or modify Core Commands methods?** [ ] Yes [ ] No

If YES, verify all steps completed:

- [ ] Updated the annotated Core Commands interface and implementation
- [ ] Built Release so source generators refreshed Service, CLI, and MCP surfaces
- [ ] Ran `scripts\Invoke-ExcelFreeTests.ps1 -Local -Contracts`
- [ ] Verified CLI and MCP names, parameters, defaults, validation, and results match
- [ ] Updated focused integration tests for the affected entry points
- [ ] Updated canonical guidance in `skills/shared` and user documentation when behavior changed

**Coverage Impact**: +___ methods, ___% → ___% coverage

## Checklist
- [ ] Code follows project style guidelines
- [ ] Self-review of code completed
- [ ] Code builds with zero warnings
- [ ] Appropriate error handling added
- [ ] Updated help text (if adding new commands)
- [ ] Updated README.md (if needed)
- [ ] Follows AGENTS.md and the applicable COM safety guide
- [ ] Uses batch API with proper disposal (`using var batch` or `await using var batch`)
- [ ] Properly handles 1-based Excel indexing
- [ ] Escapes user input with `.EscapeMarkup()`
- [ ] Returns consistent exit codes (0 = success, 1+ = error)

## Additional Notes
Any additional information that reviewers should know.
