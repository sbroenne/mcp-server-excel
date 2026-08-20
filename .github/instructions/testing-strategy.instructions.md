---
applyTo: "tests/**/*.cs"
excludeAgent: "code-review"
---

# Testing strategy

Use the narrowest existing project and filter that cover the change. Set an
explicit timeout in the terminal/tool invocation for every Excel-dependent run.

## Commands

```powershell
# One Core feature
dotnet test tests\ExcelMcp.Core.Tests\ExcelMcp.Core.Tests.csproj --filter "Feature=PowerQuery&RunType!=OnDemand"

# One MCP tool
dotnet test tests\ExcelMcp.McpServer.Tests\ExcelMcp.McpServer.Tests.csproj --filter "FullyQualifiedName~PowerQuery"

# One CLI command group
dotnet test tests\ExcelMcp.CLI.Tests\ExcelMcp.CLI.Tests.csproj --filter "FullyQualifiedName~PowerQuery"

# Session and batch infrastructure
dotnet test tests\ExcelMcp.ComInterop.Tests\ExcelMcp.ComInterop.Tests.csproj --filter "RunType=OnDemand"

# One named test while debugging
dotnet test tests\ExcelMcp.Core.Tests\ExcelMcp.Core.Tests.csproj --filter "FullyQualifiedName=Namespace.Class.Method"
```

VBA tests require Excel Trust Center access. Screenshot tests run separately
because they use desktop/clipboard resources.

## Test design

- COM-dependent behavior uses real Excel integration tests. Do not mock
  `IExcelBatch` or Excel COM to claim interop coverage.
- Pure parser, mapper, generator, and serialization behavior may use fast
  non-COM tests when no Excel behavior is involved.
- A regression test must fail for the missing behavior before the implementation
  changes and pass afterward.
- Verify the resulting workbook state and all relevant result fields, not only
  `Success`.
- Test errors with precise assertions; do not accept multiple incompatible
  outcomes.

## Isolation and fixtures

- Create a unique workbook per test with the established fixture/helper.
- Do not share mutable workbooks between tests.
- Do not combine `IClassFixture<T>` with a collection fixture on the same class;
  that can create competing Excel sessions.
- Use `.xlsm` for VBA scenarios.
- Include the repository's required traits (`Category`, `Feature`, `Layer`,
  `RequiresExcel`, `Speed`, and `RunType` where applicable) by following nearby
  tests in the same feature.

## Save and round-trip behavior

- Do not call `batch.Save()` for in-memory assertions.
- Save only when persistence is the behavior under test.
- For persistence, save/close, reopen in a new batch, and assert the state again.
- For update/replace operations, assert old content is absent and new content is
  exact; a successful operation result does not prove replacement semantics.

## Debugging failures

Run the failing test alone first. Check workbook isolation, fixture choice,
actual Excel state, COM cleanup, and persistence assumptions before broadening
the run. Do not hide a failure with skip/xfail or weaken a deterministic
assertion.
