# ExcelMcp repository rules

Windows and PowerShell; use the SDK selected by `global.json`. Desktop Excel
is required for COM tests. GitHub-hosted runners do not have Excel.

MCP Server and `excelcli` are equal entry points: behavior, defaults, validation,
results, and documentation must agree. Read `CONTEXT.md` for the system map and
only the `.github/instructions` files matching the work. For extension changes,
also read `vscode-extension/.github/instructions/extension-development.instructions.md`.

## Implementation

- Core `[ServiceCategory]` interfaces drive generated Service, CLI, and MCP
  routing. Change contracts/generators, not emitted code. Follow a changed
  contract through both entry points, tests, and shared guidance.
- Behavioral changes require a focused failing regression test before the fix.
  Documentation/configuration-only changes do not need synthetic tests.
- `Success == true` requires an empty or null `ErrorMessage`.
- Keep customer/workbook data, credentials, connection strings, and private
  paths out of public artifacts. Keep temporary notes outside the repository.

## Build and validation

```powershell
dotnet restore Sbroenne.ExcelMcp.sln
dotnet build Sbroenne.ExcelMcp.sln -c Release --no-restore
```

Build with zero warnings. Use targeted tests; see
[testing strategy](instructions/testing-strategy.instructions.md).
Runtime changes in Core, ComInterop, Service, CLI, MCP, or their generators also
require `scripts\Test-E2E.ps1` locally with Excel. Report it as not run when
Excel is unavailable; build-only checks do not cover COM.

Run applicable existing checks, not replacement audits:

```powershell
& .\scripts\check-com-leaks.ps1
& .\scripts\audit-core-coverage.ps1 -CheckNaming -FailOnGaps
& .\scripts\check-mcp-core-implementations.ps1
& .\scripts\check-success-flag.ps1
& .\scripts\check-doc-counts.ps1 -SkipBuild
& .\scripts\check-dynamic-casts.ps1
```

`-SkipBuild` requires a successful Release solution build in this worktree.
Otherwise omit it. PRs record the root cause, affected contracts, and validation.

## Git and release

- Never commit to `main`, force-push, or bypass hooks. Report hook blockers.
- Coding-agent assignments requesting repository changes authorize delivery
  commits and a PR. Otherwise ask before commit/push. Merging and publishing
  require separate authorization.
- User-visible changes require a changeset; internal/docs/tests/CI changes use
  the `skip-changelog` PR label. Versions and `CHANGELOG.md` are release-generated.
- Plugin publication changes must follow
  `.github/workflows/docs/publish-plugins-setup.md#maintenance-and-updates`;
  the published repository is output-only.
