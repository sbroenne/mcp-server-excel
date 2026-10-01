# ExcelMcp repository rules

Windows and Apple Silicon macOS; use PowerShell 7 and the SDK selected by
`global.json`. macOS support is experimental beta with capability gates.
Desktop Excel is required for platform-dependent tests; COM tests need Windows.
GitHub-hosted runners do not have Excel.

MCP Server and `excelcli` are equal entry points: behavior, defaults, validation,
results, and documentation must agree. Read `CONTEXT.md` for the system map and
only the `.github/instructions` files matching the work. For extension changes,
also read `vscode-extension/.github/instructions/extension-development.instructions.md`.

Every action marked supported on macOS must match the Windows Excel/public
contract 1:1: inputs, defaults, validation, results, workbook effects,
persistence, and errors. macOS may use a different supported Excel API and has
different shared-application ownership, but it must not silently narrow,
approximate, or partially implement a supported action. Gate the whole action
or an explicitly declared unsupported variant when exact parity is impossible.

## Implementation

- Core `[ServiceCategory]` interfaces drive generated Service, CLI, and MCP
  routing. Change contracts/generators, not emitted code. Follow a changed
  contract through both entry points, tests, and shared guidance.
- Keep agent-facing metadata, server instructions, skills, and recovery messages
  consistent with implemented behavior. Follow `mcp-llm-guidance.instructions.md`
  for guidance changes; do not edit generated or installed skill copies.
- Preserve exact tool, action, parameter, and flag names in agent guidance.
  When MCP and CLI spellings differ, show both explicitly or use native examples;
  do not replace identifiers with vague descriptions.
- Do not tell agents to create or retain workbook backup/recovery copies, or to
  work on duplicate workbooks as a prerequisite to requested edits. Explain
  destructive consequences without adding file-copy steps. Copying or exporting
  is appropriate only when part of the user's request.
- Access workbook contents through supported Excel APIs: COM on Windows,
  Apple Events or the verified optional Office.js tier on macOS. Pre-open binary
  container detection may read only IRM/AIP protection metadata; it must not
  parse workbook content. The opaque-workbook rule below applies everywhere.
- Behavioral changes require a focused failing regression test before the fix.
  Documentation/configuration-only changes do not need synthetic tests.
- `Success == true` requires an empty or null `ErrorMessage`.
- MCP Server and `excelcli` are the only supported product entry points. Do not
  preserve or add public Core/ComInterop compatibility APIs for hypothetical
  external consumers when neither entry point uses them.
- Treat workbook files as opaque. Never create, parse, inspect, or mutate ZIP,
  OOXML, relationship, custom XML, or DataMashup internals in production code,
  tests, scripts, or fixtures, including through a package/Open XML library.
  Use Excel-supported APIs or a trusted helper that automates Excel; otherwise
  report the capability as unsupported. Copying an intact Excel-authored
  workbook or template as an opaque whole file is allowed.
- Keep customer/workbook data, credentials, connection strings, and private
  paths out of public artifacts. Keep temporary notes outside the repository.

## Build and validation

```powershell
dotnet restore Sbroenne.ExcelMcp.sln
dotnet build Sbroenne.ExcelMcp.sln -c Release --no-restore
```

Build with zero warnings. Use targeted tests; see
[testing strategy](instructions/testing-strategy.instructions.md).
Run every Excel-dependent test command sequentially; never overlap Excel test
fixtures, test hosts, or E2E runs.
Runtime changes in Core, ComInterop, Service, CLI, MCP, or their generators also
require local desktop Excel E2E: `scripts\Test-E2E.ps1` on Windows and
`scripts/Test-MacE2E.ps1` on macOS. Report unavailable platform execution as
not run; a cross-target build does not cover Excel behavior.

Run applicable existing checks, not replacement audits:

```powershell
& .\scripts\check-com-leaks.ps1
& .\scripts\Invoke-ExcelFreeTests.ps1 -Local -Contracts
& .\scripts\check-success-flag.ps1
& .\scripts\check-doc-counts.ps1 -SkipBuild
& .\scripts\check-dynamic-casts.ps1
& .\scripts\check-workbook-package-access.ps1
```

`-SkipBuild` requires a successful Release solution build in this worktree.
Otherwise omit it. PRs record the root cause, affected contracts, and validation.

## Git and release

- Never commit directly or force-push to `main`.
- Never skip, disable, suppress, or bypass a Git hook for any reason, including
  transient failures or previously passing validation. Fix the failure or
  report the blocker, then rerun the normal hooked command.
- Rewriting a feature branch's remote history requires explicit user
  authorization. Use `--force-with-lease` with the expected remote commit;
  never use plain `--force`. If the lease fails, stop and inspect the remote
  changes before retrying.
- Coding-agent assignments requesting repository changes authorize delivery
  commits and a PR. Otherwise ask before commit/push. Merging and publishing
  require separate authorization.
- Before finalizing a PR, resolve every review thread after addressing it, or
  dismiss it with a clear recorded reason when no change is appropriate. Never
  leave review comments unanswered or unresolved.
- User-visible changes require a changeset; internal/docs/tests/CI changes use
  the `skip-changelog` PR label. Versions and `CHANGELOG.md` are release-generated.
- Plugin publication changes must follow
  `.github/workflows/docs/publish-plugins-setup.md#maintenance-and-updates`;
  the published repository is output-only.
- Unchanged plugin output must not create publication commits/tags. Marketplace
  updates are opt-in and maintain one owned PR; follow
  `.github/workflows/docs/awesome-copilot-update-setup.md` for comparison,
  guarded writes, catch-up, and pinned workflow compilation.
