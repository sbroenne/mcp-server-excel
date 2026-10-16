# ADR-010: Share changed-area validation through SDK-native tooling

**Status:** Current

## Context and decision

Local commits and pull requests use one internal C# policy to select checks
from changed paths. Selection includes affected dependencies and preserves
test-class fixture boundaries. Explicit complete validation remains available;
unknown operational inputs fail rather than silently selecting everything.

The existing .NET SDK and xUnit own build and test execution. A small PowerShell
bootstrap loads isolated tool output, and compatibility commands forward to
the typed implementation. C# pattern safeguards run as compiler diagnostics
and banned-API checks rather than source-text scanner scripts.
This tooling is not a product entry point and does not publish distributions.

MSBuild alone interprets project files and determines restore/build dependencies
and incremental work. Do not maintain a package-consumer parser, a second
project graph, or source-content freshness stamps in the wrappers. Changed-path
routing chooses test environments, not package components or what MSBuild
must rebuild. Shared dependency files select the normal solution build and
shared checks. PR CI runs the normal npm and complete package checks without
changed-path package selection. Explicit manual commands may name components.
Package-consumer parsing and package-selection tests are deliberately removed.

The pre-commit hook uses the selected Release build and existing `Fast` unit-test
group's CLI, MCP Server, Core, and Service tests only, exposed as `Commit`.
ComInterop-specific tests remain in normal CI, with real-Excel cases requiring
a Windows/Excel environment. The hook does not run the `Process` or `Tooling` groups, package preparation,
Excel tests, or acceptance. Those checks remain required for affected changes
before PR delivery, using hosted CI where the environment supports them.
This keeps a commit useful as a small checkpoint rather than an hours-long
integration run; passing the hook does not establish complete PR validation.

## Reasons and tradeoffs

Independent script policies drift, while blanket test runs obscure which
behavior a change affects. A saved selection provides a common contract for
local and hosted checks and records why each owner was selected.

A new build framework would add another dependency without replacing the
repository's standard build and test tools. Moving decisions into C# makes
them directly testable while keeping Windows and existing callers supported.
Isolated bootstrap output also avoids locking the executing tool during
rebuild tests; MSBuild owns compilation freshness.

Source-derived test ownership is not a substitute for actual test execution.
Selected reports must establish that cases ran and passed; compiler diagnostics
cannot establish Excel behavior. Supported test-classification rules are checked
at compilation, while built-assembly inventories and feature-group tests still
validate runtime selection semantics. Shared inputs can legitimately select
several owners, and ownership needs focused regressions as the test structure evolves.

Changed-area selection, native CLI workflow checks, package
preparation, and result verification are xUnit cases or typed components.
Some PowerShell scripts still own work the typed tool calls, such as the
explicit complete acceptance command and cloud runner entry point
(`Test-E2E.ps1`, with its own stage filters and OLAP setup) and identity-checked
CLI service cleanup. Documentation-count validation instead runs during MCP
compilation against documentation inputs and the compiled Core manifest.
Compatibility commands kept for existing callers forward to the typed tool or to those
scripts rather than holding their own selection policy. Standard SDK, npm, and
VSIX tools still produce the distributions.

Compiler checks report supported semantic patterns earlier and avoid a second
regex policy. They cannot prove correct HRESULT classification, complete COM
lifetime handling, or actual workbook state. Runtime tests and reviewer judgment
remain necessary; executable scripts that prepare environments, orchestrate
Excel, or validate prepared outputs are not all replaceable by compilation.

## Implementation and guidance

- [Selection policy](../tools/ExcelMcp.Build/ValidationPolicy.cs)
- [Execution and result verification](../tools/ExcelMcp.Build/TestExecution.cs)
- [Compiler safeguards](../src/ExcelMcp.Analyzers/) and [build integration](../Directory.Build.targets)
- [Package preparation](../tools/ExcelMcp.Build/PackageExecution.cs)
- [Built Excel group inventory](../tools/ExcelMcp.Build/ExcelGroupExecution.cs)
- [Build constraints](agents/rules/development-workflow.md)
- [Commands and test boundaries](../tests/README.md)
