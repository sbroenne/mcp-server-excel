# Pre-commit checks

The hook validates the proposed commit. It never prepares release artifacts for
publication, installs packaging dependencies, automatically stages generated
files, or stashes your work. Packaging regression tests use temporary fixtures;
required PR checks build and inspect the actual affected release packages.

## Install and run

From the repository root, create the Git hook wrapper using PowerShell 7:

```powershell
$hook = git rev-parse --git-path hooks/pre-commit
@'
#!/bin/sh
exec pwsh -NoProfile -File scripts/pre-commit.ps1
'@ | Set-Content -LiteralPath $hook -Encoding utf8NoBOM
```

This uses Git's resolved hooks directory, including in a worktree. If your
installation has a custom `core.hooksPath`, install the wrapper there instead.
Do not replace an existing hook without preserving its other checks.

Run the same checks manually:

```powershell
& .\scripts\pre-commit.ps1
```

## What runs locally

Every commit is checked for direct commits to `main` and nonportable staged
npm lockfiles. Further checks are selected from staged paths by
`scripts\Get-ValidationPlan.ps1`.

| Changed inputs | Release build | Local Excel E2E | Release artifact creation |
|---|---|---|---|
| Documentation, website, videos, Azure or analytics maintenance without script regressions | No | No | Never |
| Extension, npm wrappers, or Claude bundle inputs without packaging regressions | No | No | Never |
| Packaging/release scripts, tested Azure scripts, or evidence capture script | Yes | No | Never |
| Tests, hook, CI policy, analyzer settings, or skill templates | Yes | No | Never |
| Core, COM, Service, Cleanup, either entry point, runtime generators | Yes | Yes | Never |
| Central build/dependency settings or embedded shared skill guidance | Yes | Yes | Never |
| Unknown inputs | Yes | Yes | Never |

Mixed commits combine the applicable checks. During a merge, selection compares
against the incoming parent so imported changes do not trigger unrelated work.
The hook prints its reasons before running expensive checks.

Runtime changes run the retained source-pattern guards, a Release build,
focused generated-contract checks, and the complete `Test-E2E.ps1` sequence.
The guards flag suspicious patterns; they do not prove every COM lifetime or
error-result path is correct. Local build cleanup remains pipe-scoped and must
not stop another worktree's sessions.

Changed test projects run their normal Excel-free tests. Skill preparation,
packaging/release, and script safety checks live in separate test projects.
Packaging changes select their regression tests, which may build packages in
temporary folders and use disposable local publication repositories. Actual
release artifact creation belongs to CI or explicit manual validation.
Feature-specific Excel tests remain the author's responsibility; the E2E
sequence is not the full feature suite.

## Failures and partial staging

Before build-based checks, relevant source, test, skill, script and central
configuration inputs must match the Git index. Unstaged or untracked inputs
that would change the build stop validation. Stage or set aside those changes
explicitly; the hook never changes the index for you. Unrelated documentation
edits are left alone.

A failing command stops the hook and retains its output and exit code.
Windows with desktop Excel is required when Excel checks are selected.
An unavailable prerequisite is not reported as a pass. Report the blocker
rather than bypassing the hook.

Fix generated-contract failures in the annotated interfaces or generators,
not in generated enum or adapter files.

## PR and manual package checks

`CI Gate` remains present on every PR. It builds Release, runs normal
`RequiresExcel=false` selections sequentially with deadlines, runs the source
guards and npm launcher tests, and creates and checks the affected packages.
`Docs Site` builds and audits the website separately.

GitHub-hosted runners do not have Excel. Record local E2E and feature-specific
results in the PR; CI does not replace them.

Explicit manual package validation uses the same command as PR CI:

```powershell
& .\scripts\Build-ReleasePackages.ps1
& .\scripts\Build-ReleasePackages.ps1 -Components Extension,Mcpb
```

The command writes a new directory under `artifacts\packages`, checks installed
npm launchers and archive contents, and never publishes packages or creates
tags. The extension reuses prepared MCP executables. The Claude bundle copies
metadata only and configures npx to resolve the npm server at launch; MCPB-only
packaging does not publish a runtime.

Isolated hook regressions:

```powershell
dotnet test tests\ExcelMcp.ScriptSafety.Tests\ExcelMcp.ScriptSafety.Tests.csproj -c Release --filter "Feature=PreCommit" --blame-hang-timeout 60s
```
