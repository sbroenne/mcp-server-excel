# ADR-010: Share changed-area validation through SDK-native tooling

**Status:** Current

## Context and decision

Local commits and pull requests share explicit changed-path rules in
`scripts/Get-ValidationPlan.ps1`. Saved plans record owning groups, feature
filters and affected dependencies. Unknown operational inputs fail rather than
silently selecting everything. Git comparisons preserve merge-base selection
for pull requests and merge-parent selection for staged commits.

The .NET SDK and xUnit own project evaluation, restore, incremental compilation,
test discovery and test execution. Shared PowerShell coordinates these commands;
there is no custom C# build executable, bootstrap cache or options-JSON
dispatcher. Selection and CI completion do not need the .NET SDK.
Compiler diagnostics and banned-API checks retain supported C# safeguards.
The scripts are development automation, not product entry points.

MSBuild alone evaluates project dependencies. Selected multi-project builds use
a temporary solution filter, not a second project graph.
Owned builds disable reusable MSBuild nodes and the persistent compiler server,
while retaining normal incremental compilation.
Shared dependency inputs select a complete solution build. The first selected CI group in
Fast/Process/Tooling order owns that build; other groups build their own projects.
PR CI always runs normal npm checks and complete package checks. Changed paths
do not select package components; explicit manual commands may do so.

## Reasons and tradeoffs

The previous C# policy made checks testable but duplicated responsibilities
already available in `dotnet test` and xUnit. It also required building a tool
before merely selecting checks or verifying Actions results. A local spike
confirmed that native settings reject empty runs and skipped cases, report
fixture-cleanup failures, and support the existing feature filters.

The source-derived catalogue is removed, not translated into PowerShell.
Command implementations select explicit native feature labels and known
dependencies. Test edits select their owning project group; unit and recording-only
protocol tests remain Excel-free. Real-Excel MCP transport fixtures, their shared
helpers and other test folders can select that owner's Excel cases. Edits to
partial `PersistentServiceWindowTests` files select the complete owning test
class, preserving its fixture without selecting unrelated Service Excel classes.
Shared runtime/helper changes intentionally select broader runtime,
acceptance and infrastructure checks. This trades some precision for rules that
are smaller and easier to review. New areas and changed label ownership need
explicit rules and regression coverage.

The commit hook remains narrow: selected Release builds and quick CLI, MCP
Server, Core and Service tests (`Commit` projected from `Fast`), excluding
`AdapterTestKind=System` and `Speed=Slow`. It does not run process/tooling
groups, packaging, Excel or acceptance. CI retains excluded Excel-free coverage.
Run local process checks when that boundary changes, not on every push.
Selected Excel checks still require local desktop Excel; hosted runners cannot
replace them. A passing hook is not complete PR validation.

### Native execution and evidence

`tests/strict.runsettings` enables `TreatNoTestsAsError`, `FailSkips` and
`FailWarns`. Stages add native session and hang limits, fresh TRX reports,
diagnostics and ownership directories. The existing xUnit architecture test
exports the built Excel inventory and preserves class-fixture boundaries.
Excel and acceptance runs reconcile discovered and executed test names,
including theory-row multiplicities. A successful exit alone does not establish
that all required cases ran.

A native test timeout stops the test host but can leave a child process alive,
even holding redirected output open after its parent exits. The small shared
PowerShell process helper therefore uses a Windows job object. A suspended command is
assigned before it can run; the operating system retains descendant
ownership after intermediate parents exit. The deadline covers the command,
descendants and output drain. The few Windows API bindings are loaded inside
PowerShell; they are not another executable, build framework or SDK bootstrap.
Cleanup is limited to the owned job and exact retained process handles.
The standard .NET process-tree termination also covers packaged PowerShell
activations that Windows can launch outside the job.
CLI service cleanup continues to use its existing pipe/identity boundary.
Primary and cleanup failures must both remain visible.

### Packaging ownership

`Build-ReleasePackages.ps1` coordinates the existing named PowerShell package
commands directly. Shared helpers cover safe filesystem/archive handling,
deadlines and output inspection, without another dispatcher or build graph.
Actions and optional local troubleshooting use the same scripts; availability
locally does not make packaging a commit prerequisite. Actions owns authorized
publication and reruns checks on submitted source.

Generated skills are officially validated before hosted package construction.
Extension source checks precede runtime publishing; one compilation serves
both VSIX targets. Finished-package inspection and native runtime checks remain
necessary afterward. Coordinated metadata and prepared-output ownership remain
as recorded in [ADR-009](ADR-009-COORDINATED-RELEASE-OUTPUTS.md).

This revision replaces the earlier C# selection/execution choice and its later
packaging-only revision. PowerShell's dynamic scope and native exit handling
still require behavioral tests. Compiler checks cannot prove correct HRESULT
classification, complete COM ownership or actual workbook state.

## Implementation and guidance

- [Selection policy](../scripts/Get-ValidationPlan.ps1)
- [Native coordination](../scripts/ValidationExecution.ps1)
- [Stage evidence](../scripts/Invoke-TestStage.ps1) and [native settings](../tests/strict.runsettings)
- [Owned command execution](../scripts/Invoke-NativeCommand.ps1)
- [Compiler safeguards](../src/ExcelMcp.Analyzers/) and [build integration](../Directory.Build.targets)
- [Package preparation](../scripts/Build-ReleasePackages.ps1)
- [Built Excel inventory](../scripts/Get-ExcelTestGroups.ps1)
- [Build constraints](../scripts/AGENTS.md)
- [Commands and test boundaries](../tests/README.md)
