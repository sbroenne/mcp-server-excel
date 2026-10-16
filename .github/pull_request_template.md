## Summary
Explain the problem, root cause, and intended outcome. Link related issues
(`Closes #...`, `Relates to #...`) or write N/A.

## Changes and consequences
Describe affected behavior and contracts, migration or partial-state/recovery
consequences, and documentation changes. MCP Server and `excelcli` are equal
entry points: explain any intentional difference.
For contract/generator changes, follow
[Core guidance](https://github.com/sbroenne/mcp-server-excel/blob/main/src/ExcelMcp.Core/AGENTS.md)
and record the generated surfaces inspected.

## Evidence
Follow [AGENTS.md](https://github.com/sbroenne/mcp-server-excel/blob/main/AGENTS.md)
and the applicable nested guides. List each exact command and its result
(passed, failed, or not run); include the changed-area plan and explain gaps.
For behavior fixes, identify the regression that failed before the fix and what
returned fields or actual workbook state it verifies.

Desktop Excel is required for COM/acceptance checks; ordinary hosted runners
cannot supply that evidence. Run Excel commands sequentially and preserve owned
process/session cleanup. Compilation is not proof of COM ownership or native
behavior. Python evaluations are on-demand only.

## Review and release
Record independent complete-branch review findings and their disposition,
followed by `ponytail-review` when available. Address and resolve review threads
or record why no change is appropriate. Never bypass Git hooks.

Link the [changeset](https://github.com/sbroenne/mcp-server-excel/blob/main/.changeset/README.md)
for user-visible changes, or explain why this is internal/docs/tests/CI-only
and uses `skip-changelog`. Versions and changelogs are release-generated.

## Limitations
Record remaining risks, unavailable validation, and any screenshots that
actually help explain a visual change (synthetic data only).
Do not include customer/workbook data, credentials, connection strings, or
private paths. Merging and publishing require separate authorization.
