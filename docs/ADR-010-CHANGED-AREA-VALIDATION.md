# ADR-010: Share changed-area validation through SDK-native tooling

**Status:** Current

## Context and decision

Local commits and pull requests use one internal C# policy to select checks
from changed paths. Selection includes affected dependencies and preserves
test-class fixture boundaries. Explicit complete validation remains available;
unknown operational inputs fail rather than silently selecting everything.

The existing .NET SDK and xUnit own build and test execution. A small PowerShell
bootstrap loads isolated tool output, and compatibility commands forward to
the typed implementation. Source safeguards run during ordinary production
compilation, not as a separate source scanner.
This tooling is not a product entry point and does not publish distributions.

## Reasons and tradeoffs

Independent script policies drift, while blanket test runs obscure which
behavior a change affects. A saved selection provides a common contract for
local and hosted checks and records why each owner was selected.

A new build framework would add another dependency without replacing the
repository's standard build and test tools. Moving decisions into C# makes
them directly testable while keeping Windows and existing callers supported.
Isolated bootstrap output also avoids locking the executing tool during
rebuild tests.

Source-derived test ownership is not a substitute for actual test execution.
Selected reports must establish that cases ran and passed; compilation and
source guards cannot establish Excel behavior. Shared inputs can legitimately
select several owners, and source ownership needs focused regressions as the
test structure evolves.

Changed-area selection, native CLI workflow checks, source safety regressions, package
preparation, and result verification are xUnit cases or typed components.
Some PowerShell scripts still own work the typed tool calls, such as the
explicit complete acceptance command and cloud runner entry point
(`Test-E2E.ps1`, with its own stage filters and OLAP setup), the documentation
count validator, and identity-checked CLI service cleanup. Compatibility
commands kept for existing callers forward to the typed tool or to those
scripts rather than holding their own selection policy. Standard SDK, npm, and
VSIX tools still produce the distributions.

## Implementation and guidance

- [Selection policy](../tools/ExcelMcp.Build/ValidationPolicy.cs)
- [Execution and result verification](../tools/ExcelMcp.Build/TestExecution.cs)
- [Compiler source safeguards and scope](SOURCE-SAFETY.md)
- [Internal compiler analyzer](../src/ExcelMcp.Analyzers/ExcelSafetyAnalyzer.cs)
- [Package preparation](../tools/ExcelMcp.Build/PackageExecution.cs)
- [Built Excel group inventory](../tools/ExcelMcp.Build/ExcelGroupExecution.cs)
- [Build constraints](agents/rules/development-workflow.md)
- [Commands and test boundaries](../tests/README.md)
