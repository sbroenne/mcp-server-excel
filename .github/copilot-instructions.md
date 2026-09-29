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
- Production code must access workbook contents through Excel COM. Never open
  an Excel file as a ZIP/OOXML package or parse/modify its internal XML outside
  tests. Pre-open binary container detection may read only IRM/AIP protection
  metadata; it must not parse workbook content.
- Behavioral changes require a focused failing regression test before the fix.
  Documentation/configuration-only changes do not need synthetic tests.
- `Success == true` requires an empty or null `ErrorMessage`.
- MCP Server and `excelcli` are the only supported product entry points. Do not
  preserve or add public Core/ComInterop compatibility APIs for hypothetical
  external consumers when neither entry point uses them.
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
& .\scripts\check-doc-counts.ps1 -SkipBuild -AllowStaleAdvertisedCounts
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
