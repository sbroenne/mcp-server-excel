# ADR-010: Share changed-area validation through SDK-native tooling

**Status:** Current

## Context and decision

Local commits and pull requests use one internal C# policy to select checks
from changed paths. Selection includes affected dependencies and preserves
test-class fixture boundaries. Explicit complete validation remains available;
unknown operational inputs fail rather than silently selecting everything.

The existing .NET SDK and xUnit own build and test execution. A small PowerShell
bootstrap loads isolated tool output, and compatibility commands forward to
the typed implementation. Source safeguards live in the same internal tool.
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

Native CLI workflow checks are independently selectable xUnit cases rather than
a nested PowerShell test runner. Package preparation, source checks, test stages,
built test inventory, CI build inputs and completion checks use typed components.
Retained PowerShell commands preserve existing callers and translate arguments;
they do not maintain a second selection or result-checking policy. Standard
SDK, npm and VSIX tools still produce the distributions. Cloud Excel runner
operations and safety checks remain unchanged.

## Implementation and guidance

- [Selection policy](../tools/ExcelMcp.Build/ValidationPolicy.cs)
- [Execution and result verification](../tools/ExcelMcp.Build/TestExecution.cs)
- [Source safeguards](../tools/ExcelMcp.Build/SourceGuards.cs)
- [Package preparation](../tools/ExcelMcp.Build/PackageExecution.cs)
- [Built Excel group inventory](../tools/ExcelMcp.Build/ExcelGroupExecution.cs)
- [Build constraints](agents/rules/development-workflow.md)
- [Commands and test boundaries](../tests/README.md)
