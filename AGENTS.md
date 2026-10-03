# ExcelMcp agent instructions

ExcelMcp automates installed desktop Excel through COM on Windows. Use
PowerShell and the SDK selected by `global.json`. Desktop Excel is required
for COM tests; GitHub-hosted runners do not have Excel.

MCP Server and `excelcli` are equal entry points: behavior, defaults, validation,
results, and documentation must agree.

For an unfamiliar area, start with [CONTEXT.md](CONTEXT.md) for the system map
and terminology. [Architecture decisions](docs/DECISIONS.md) explain current
choices and tradeoffs; read the records relevant to an architectural change,
not the entire collection for every task. Surface conflicts before changing a
decision. Actionable rules belong here and in the applicable guides below.

Work requests are tracked in GitHub Issues; see
[issue tracker](docs/agents/issue-tracker.md). Contributor setup and client
instruction discovery are in [agent development](docs/agents/development.md).

## Task-specific guidance

Before changing files, read the matching guides below, including nested
`AGENTS.md` files. Native discovery differs between Copilot, Claude Code, and
Codex; launching at the root does not guarantee every nested file is loaded.
Paths are relative to the repository root.

| Work | Required guidance |
| --- | --- |
| `src/**/*.cs` | [Runtime boundaries](docs/agents/rules/architecture-patterns.md) |
| Core commands/action models, Service, CLI, MCP, or generators | [Generated contracts](docs/agents/rules/coverage-prevention-strategy.md) |
| Core or ComInterop C# | [COM safety](docs/agents/rules/excel-com-interop.md) |
| Connection commands/sanitizer or connection tests/helpers/fixtures | [Connections](docs/agents/rules/excel-connection-types-guide.md) |
| `tests/**/*.cs` | [Testing strategy](tests/AGENTS.md) |
| Workflows, project files, SDK/build configuration, PowerShell scripts | [Build and release](docs/agents/rules/development-workflow.md) |
| README/index files, FEATURES, CHANGELOG, SECURITY, PRIVACY, docs, specs, skills, or website | [Documentation](docs/agents/rules/documentation-structure.md) |
| Skill sources, Core command metadata, generators, MCP, or Build-AgentSkills.ps1 | [Product agent guidance](docs/agents/rules/mcp-llm-guidance.md) |
| MCP Server or MCP generator | [MCP boundaries](docs/agents/rules/mcp-server-guide.md) |
| `llm-tests/**` | [LLM evaluations](llm-tests/AGENTS.md) |
| Repository/agent instructions, CONTEXT, or `docs/agents/**` | [Instruction maintenance](docs/agents/rules/meta.md) |
| `vscode-extension/**` | [Extension](vscode-extension/AGENTS.md) |
| `videos/excel-mcp-intro/**` | [Video](videos/excel-mcp-intro/AGENTS.md) |

## Implementation

- Core `[ServiceCategory]` interfaces drive generated Service, CLI, and MCP
  routing. Change contracts/generators, not emitted code. Follow a changed
  contract through both entry points, tests, and shared guidance.
- Keep agent-facing metadata, server instructions, skills, and recovery messages
  consistent with implemented behavior. Follow the product agent guidance above;
  do not edit generated or installed skill copies.
- Preserve exact tool, action, parameter, and flag names in agent guidance.
  When MCP and CLI spellings differ, show both explicitly or use native examples;
  do not replace identifiers with vague descriptions.
- Do not tell agents to create or retain workbook backup/recovery copies, or to
  work on duplicate workbooks as a prerequisite to requested edits. Explain
  destructive consequences without adding file-copy steps. Copying or exporting
  is appropriate only when part of the user's request.
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

Build with zero warnings. Use targeted tests; see the
[testing strategy](tests/AGENTS.md). Run every Excel-dependent test command
sequentially; never overlap Excel test fixtures, test hosts, or E2E runs.
Runtime changes in Core, ComInterop, Service, CLI, MCP, or their generators also
require `scripts\Test-E2E.ps1` locally with Excel. Report it as not run when
Excel is unavailable; build-only checks do not cover COM.

The Python evaluation suite under `llm-tests/` is on-demand only, not part of the
normal development lifecycle. This includes offline checks, Excel fixture checks,
SDK discovery, and live agent comparisons. Run it and install its dependencies
only when the user explicitly requests evaluation work. Unrun evaluations or
unavailable Python/SDK dependencies are not implementation, commit, PR, or merge blockers.
This does not waive required .NET/COM/E2E validation or normal Git hooks.

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

## Code Review Rules

Report high-confidence defects introduced by the change, not style or unrelated
cleanup. Read the applicable [review checks](docs/agents/review.md).
Review tasks do not require installing dependencies, running builds, or editing
code unless requested.

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
