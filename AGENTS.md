# ExcelMcp agent instructions

ExcelMcp automates installed desktop Excel through COM on Windows. Use
PowerShell 7.4 or later (`pwsh`) and the SDK selected by `global.json`. Desktop Excel is required
for COM tests; GitHub-hosted runners do not have Excel.

MCP Server and `excelcli` are equal entry points: behavior, defaults, validation,
results, and documentation must agree.

ExcelMcp always operates in the current user's Windows session and security
context. Discover and reconnect only to Excel instances accessible in that
context; do not seek access to other users' sessions or elevate permissions.
Workbook location does not change this scope. Running in the user's context
does not authorize changes to unrelated workbooks.

## Code Review Rules

Report high-confidence defects introduced by the change, not style or unrelated
cleanup. Establish a reachable failure, its affected caller/workbook state, and
the missing or misleading test evidence. Review does not require builds,
dependency installation, or edits unless requested.

- Trace changed catches and fallbacks through their caller. Logging an error
  then returning empty/default data, skipping a failed item, or reporting success
  still hides failure. Optional COM reads may suppress only exact documented
  unavailable-member HRESULTs; `0x800A03EC` is not a missing-property diagnosis.
  Unexpected COM errors must fail with their context.
- Check failure ownership: Core leaves execution errors to batch/Service;
  transport boundaries may convert failures, not erase them. Preserve original
  cause, cancellation, timeout, session recovery, and exact partial state.
  Cancellation is not rollback; reject unrequested workbook cleanup.
- Review COM ownership across every exit and failure path. Acquired references
  need `finally` release, not session-owned objects. Process cleanup needs
  PID/start-time identity and retained handles. EXL011 and EXL005 reject dynamic
  declarations and late-bound calls in production and test builds.
- Follow changed source contracts through generated Service, CLI options/batch
  JSON, MCP schemas, and manual exceptions. Names, defaults, validation, results,
  and timeouts must agree. Adapters must not acquire Core/session ownership.
- Check MCP protocol output, including bootstrap paths, not just direct console
  calls. Telemetry must never wait on setup, export, flush, or disposal on product
  paths, delay incremental results/cancellation/exit, or alter results/errors.
  Bound queues; verify local delivery separately from responsiveness and privacy.
- Check guidance against implemented inputs, defaults, and capabilities.
  Distinguish MCP snake_case inputs, CLI flags, nested JSON keys, and outputs.
  Flag nonexistent actions, forced unrequested work, emojis, required unattended
  screenshots, and consent prose presented as implemented server elicitation.
- Tests must establish returned fields and actual Excel state, including the
  partial state after failure; success alone or mocked COM is not evidence.
  Compiler checks catch supported patterns, not correct HRESULT semantics,
  complete COM ownership, runtime state, or matching entry-point behavior.
- Use `ponytail-review` when available after correctness/security review for
  concrete behavior-preserving simplifications. Required COM cleanup,
  validation, failure reporting, tests, and generated contracts are not bloat.

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
| `src/**/*.cs` | [Runtime boundaries](src/AGENTS.md) |
| Core commands/action models, Service, CLI, MCP, generators, or `.github/usage-analytics-weights.json` | [Generated contracts](src/ExcelMcp.Core/AGENTS.md) |
| Core or ComInterop C# | [COM safety](src/ExcelMcp.ComInterop/AGENTS.md) |
| Connection commands/sanitizer or connection tests/helpers/fixtures | [Connections](src/ExcelMcp.Core/Commands/Connection/AGENTS.md) |
| `tests/**/*.cs` | [Testing strategy](tests/AGENTS.md) |
| Workflows, project files, SDK/build configuration, PowerShell scripts, `.changeset/**` | [Build and release](scripts/AGENTS.md) |
| README/index files, FEATURES, CHANGELOG, SECURITY, PRIVACY, docs, specs, skills, or website | [Documentation](docs/AGENTS.md) |
| Skill sources, Core command metadata, generators, MCP, or Build-AgentSkills.ps1 | [Product agent guidance](skills/AGENTS.md) |
| MCP Server or MCP generator | [MCP boundaries](src/ExcelMcp.McpServer/AGENTS.md) |
| `llm-tests/**` | [LLM evaluations](llm-tests/AGENTS.md) |
| Repository/agent instructions, CONTEXT, or `docs/agents/**` | [Instruction maintenance](docs/AGENTS.md#instruction-maintenance) |
| `vscode-extension/**` | [Extension](vscode-extension/AGENTS.md) |
| `videos/excel-mcp-intro/**` | [Video](videos/excel-mcp-intro/AGENTS.md) |
| `videos/agentic-world-bank-briefing/**` | [Agentic workflow video](videos/agentic-world-bank-briefing/AGENTS.md) |

## Implementation

The build enforces supported COM, contract, result, dynamic-access, and test
patterns. Fix diagnostics at their source; a suppression needs a reason, not an
assumption that passing compilation proves runtime correctness.

- Use typed Excel interop in production and tests. EXL011 rejects dynamic
  declarations and late-bound operations; EXL005 rejects casts to `dynamic`.
  Comments do not exempt them. The compiler checks whether the referenced PIA
  exposes a member; runtime availability still needs real-Excel validation.
- Use Excel's native operations and validation for behavior Excel owns. Do not
  duplicate Excel's rules with custom preflight checks; add custom checks only
  for product requirements Excel does not provide.
- The user-approved document-property exception is confined to
  `ExcelDocumentProperties` in ComInterop. Its private native COM calls cover
  only workbook property collections/items, resolve member identifiers from
  Excel, and validate item identities. Do not reuse it for other features or
  add a general member-name, reflection, or dynamic fallback. Compiler checks
  remain enabled everywhere.

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
  an Excel workbook file as a ZIP/OOXML package or parse/modify its internal XML
  outside tests. The isolated Power Query user-settings archive reader is not a
  workbook-content path. Pre-open binary container detection may read only
  IRM/AIP protection metadata; it must not parse workbook content.
- Behavioral changes require a focused failing regression test before the fix.
  Documentation/configuration-only changes do not need synthetic tests.
- MCP Server and `excelcli` are the only supported product entry points. Do not
  preserve or add public Core/ComInterop compatibility APIs for hypothetical
  external consumers when neither entry point uses them.
- Keep customer/workbook data, credentials, connection strings, and private
  paths out of public artifacts. Keep temporary notes outside the repository.
- ExcelMcp acts for the current user. Task-relevant workbook contents, names,
  formulas, VBA/Power Query code, and file paths may be read and returned to
  that user through supported operations; do not hide them merely because
  they are private. This is not public disclosure or permission to access
  unrelated workbooks. Credential protections still apply; do not put private
  information in public artifacts or usage reporting.
- Returning an owned Excel dialog's visible text to its user for troubleshooting
  is allowed; this is not public disclosure. Do not copy it into public artifacts
  or telemetry, read password/input values, or answer the dialog automatically.

## Errors

- Handle failures only where the layer owns recovery or failure transport.
  Preserve the original exception and operation context. Optional reads must not
  treat arbitrary COM failures as absent data. A deliberate best-effort boundary
  must explain what failed and why continuing cannot misrepresent the operation.
  Cleanup must preserve the primary failure as well as any cleanup failure.
- Do not automatically clean up or roll back workbook changes when an
  operation fails. Report the exact partial state and error so the agent can
  inspect the workbook and decide whether cleanup is appropriate.

## Build and validation

For telemetry development, use the local Aspire dashboard first; see
[local telemetry development](docs/DEVELOPMENT.md#local-development-with-telemetry).
Verify actual logs/traces from both `excelcli` and MCP, including short-lived CLI
commands, as well as non-blocking behavior when the receiver is unavailable or
stalled. Passing in-memory tests or fast shutdown alone does not prove delivery.
If Aspire is unavailable, report the blocker; do not fall back to production.
Sending test telemetry to production Azure or changing its ingestion rules
requires explicit user approval for that specific test or deployment; available
credentials are not approval.

```powershell
dotnet restore Sbroenne.ExcelMcp.sln
dotnet build Sbroenne.ExcelMcp.sln -c Release --no-restore
```

Build with zero warnings. Use targeted tests; see the
[testing strategy](tests/AGENTS.md). Run every Excel-dependent test command
sequentially; never overlap Excel test fixtures, test hosts, or E2E runs.
You may close or stop any Excel copy on the machine, for example one left by a
test run; test code still must not. Report each copy you closed by process
number and start time.
The commit hook builds selected inputs and runs selected `Fast` unit tests from
CLI, MCP Server, Core, and Service (`Commit` group), excluding
`AdapterTestKind=System` and `Speed=Slow`. ComInterop-specific, process, tooling,
package, Excel, and acceptance checks do not run on
every commit. CI owns hosted checks; run local process checks when changing
their boundary, not on every push. A passing hook is not complete PR validation.
Local validation must follow the changed behavior, not merely the directory
containing a changed file. Start with a focused regression and the existing
tests for affected callers and failure paths. Run relevant Excel tests locally
when workbook behavior or Excel ownership changes; build-only checks do not
cover COM. Shared service, routing, or lifecycle files do not by themselves
require all feature tests, complete acceptance, or all ComInterop diagnostics.
Expand coverage when the change affects a shared contract or lifetime,
focused tests reveal wider failures, or a specific risk is not covered.
Run the full suite only when that wider scope is justified or explicitly requested.

Use `scripts\Get-ValidationPlan.ps1` to identify candidate checks, then review
the selection against the actual change. If it selects unrelated features,
run narrower existing test filters or runner groups and record the commands,
results, and reason for the narrower scope in the PR. Do not describe an
unrun broad selection as passed. This local scope decision does not waive
Git hooks or required CI checks. Unknown inputs must fail selection, not select
everything. `scripts\Test-E2E.ps1` remains the explicit complete acceptance
command; run it on a local Windows desktop with Excel.

The Python evaluation suite under `llm-tests/` is on-demand only, not part of the
normal development lifecycle. This includes offline checks, Excel fixture checks,
SDK discovery, and live agent comparisons. Run it and install its dependencies
only when the user explicitly requests evaluation work. Unrun evaluations or
unavailable Python/SDK dependencies are not implementation, commit, PR, or merge blockers.
This does not waive required .NET/COM/E2E validation or normal Git hooks.

Generate a candidate plan, review its scope, and run applicable existing checks,
not replacement audits. Run the complete plan only when its selection is relevant:

```powershell
& .\build.ps1 plan --base-ref origin/main --output artifacts\validation-plan.json
& .\build.ps1 validate --plan artifacts\validation-plan.json
```

PRs record the root cause, affected contracts, and validation.

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
- Before opening a PR or pushing further changes to one, run an independent
  review of the complete branch diff in a fresh context (for example, a
  code-review subagent) against the Code Review Rules. Fix confirmed findings
  and repeat until no high-confidence findings remain. After that review, run
  `ponytail-review` when available on the complete branch diff, following its
  scope in the Code Review Rules. Then verify that the PR description and
  changeset claim only what the diff does, and search the whole repository,
  including samples, workflows, and docs, for stale
  references to anything renamed or removed.
- Before finalizing a PR, resolve every review thread after addressing it, or
  dismiss it with a clear recorded reason when no change is appropriate. Never
  leave review comments unanswered or unresolved.
- User-visible changes require a changeset; internal/docs/tests/CI changes use
  the `skip-changelog` PR label. Versions and `CHANGELOG.md` are release-generated.
  Follow [.changeset/README.md](.changeset/README.md) for package names and
  validate the fragment against the PR's actual base before pushing.
- Plugin publication changes must follow
  `.github/workflows/docs/publish-plugins-setup.md#maintenance-and-updates`;
  the published repository is output-only.
- Unchanged plugin output must not create publication commits/tags. Awesome
  Copilot listings are edited manually using the actual published output commit;
  follow `.github/workflows/docs/publish-plugins-setup.md#awesome-copilot-listings-are-manual`.
