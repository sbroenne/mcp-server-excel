# GitHub Copilot Instructions - ExcelMcp

ExcelMcp is a Windows-only .NET 10 solution that automates desktop Excel through
COM. It has two equal user entry points: the MCP Server and `excelcli`. A feature
is incomplete if those surfaces differ in behavior, validation, defaults, or
documentation.

## Repository map

- `src/ExcelMcp.ComInterop`: STA execution, COM lifetime, sessions, shutdown
- `src/ExcelMcp.Core`: Excel operations and result models
- `src/ExcelMcp.Service`: command dispatch used by both entry points
- `src/ExcelMcp.CLI`: CLI client and daemon transport
- `src/ExcelMcp.McpServer`: MCP tools and in-process service bridge
- `src/ExcelMcp.Generators*`: generated CLI, service, and MCP surfaces
- `tests`: xUnit projects; most Core behavior tests require desktop Excel
- `skills/shared`: source of truth for shared CLI/MCP agent guidance
- `gh-pages`: MkDocs site built from canonical repository documentation

## Environment

- Use Windows and PowerShell. The pinned SDK is in `global.json`.
- Microsoft Excel is required for COM integration tests and local end-to-end
  smoke tests. GitHub-hosted runners can build the solution but do not have
  Excel; never report Excel-dependent tests as passing there.
- `.github/workflows/copilot-setup-steps.yml` prepares cloud-agent dependencies.

## Working rules

1. Read only the path-specific files under `.github/instructions` that match the
   files being changed. Those files contain the detailed implementation rules.
2. Diagnose behavior before editing. For a bug or feature, add a focused
   regression/integration test first, observe the expected failure, implement
   the fix, and rerun that test. Pure documentation and configuration changes do
   not need a synthetic test.
3. Search for the same contract or generated pattern across Core, Service, CLI,
   MCP, tests, skills, and docs. Update every affected surface.
4. Preserve unrelated worktree changes. Do not create temporary planning or
   summary files in the repository.
5. Prefer existing helpers and generated surfaces. Do not duplicate dispatch,
   validation, serialization, COM cleanup, or documentation-count logic.

## Build and validation

```powershell
dotnet restore Sbroenne.ExcelMcp.sln
dotnet build Sbroenne.ExcelMcp.sln -c Release --no-restore
```

Run the smallest test project and filter that cover the change, with an explicit
tool timeout. Do not run the full Excel integration suite during iteration.

```powershell
dotnet test tests\ExcelMcp.Core.Tests\ExcelMcp.Core.Tests.csproj --filter "Feature=PowerQuery&RunType!=OnDemand"
dotnet test tests\ExcelMcp.McpServer.Tests\ExcelMcp.McpServer.Tests.csproj --filter "FullyQualifiedName~PowerQuery"
dotnet test tests\ExcelMcp.CLI.Tests\ExcelMcp.CLI.Tests.csproj --filter "FullyQualifiedName~PowerQuery"
```

For session or batch infrastructure changes, run the targeted OnDemand tests in
`tests\ExcelMcp.ComInterop.Tests`. For Core, Service, CLI, MCP, ComInterop, or
generator changes that feed runtime surfaces, also run `scripts\Test-E2E.ps1`
locally on a machine with Excel. If the current environment has no Excel, report
that validation as not run; never imply that a build-only check covered COM.

Use the existing repository checks rather than reproducing them:

```powershell
& .\scripts\check-com-leaks.ps1
& .\scripts\audit-core-coverage.ps1 -CheckNaming -FailOnGaps
& .\scripts\check-mcp-core-implementations.ps1
& .\scripts\check-success-flag.ps1
& .\scripts\check-doc-counts.ps1
& .\scripts\check-dynamic-casts.ps1
```

## Git and release

- Never commit to `main`.
- A coding-agent assignment that explicitly asks for repository changes
  authorizes the branch commits and pull request needed to deliver that task.
  Interactive work without that authorization must ask before commit or push.
  No task implicitly authorizes merging.
- Never bypass repository hooks. If a hook fails because of the environment,
  stop and report the exact blocker.
- Add a changeset for user-visible features, fixes, or breaking changes. Do not
  edit `CHANGELOG.md` manually. Internal, documentation-only, test-only, and CI
  changes use the `skip-changelog` PR label instead.
- Keep customer names, workbook names, local paths, credentials, and other
  confidential context out of public commits, issues, and pull requests.

## Path-specific guidance

- C# architecture: `instructions/architecture-patterns.instructions.md`
- Excel COM: `instructions/excel-com-interop.instructions.md`
- Tests: `instructions/testing-strategy.instructions.md`
- MCP Server: `instructions/mcp-server-guide.instructions.md`
- Generated surface coverage: `instructions/coverage-prevention-strategy.instructions.md`
- Workflows and project files: `instructions/development-workflow.instructions.md`
- Documentation: `instructions/documentation-structure.instructions.md` and
  `instructions/readme-management.instructions.md`
- LLM evaluations: `instructions/llm-testing-philosophy.instructions.md`

Trust these instructions when they cover the task. If code and instructions
disagree, verify the current implementation and update the stale instruction in
the same change.
