---
applyTo: "tests/**/*.cs"
excludeAgent: "code-review"
---

# Testing strategy

## Commands

Select one project and feature/name filter, not the full Excel suite. Set a hard
execution timeout; returning control while a test keeps running is not a timeout.

```powershell
dotnet test tests\ExcelMcp.Core.Tests\ExcelMcp.Core.Tests.csproj --filter "Feature=PowerQuery&RunType!=OnDemand"
dotnet test tests\ExcelMcp.ComInterop.Tests\ExcelMcp.ComInterop.Tests.csproj --filter "RunType=OnDemand"
```

The second command is required for session/batch infrastructure changes; narrow
by test name where appropriate. Core OnDemand tests are optional diagnostics,
not mandatory CI gates. VBA needs Trust Center access; run screenshots
separately because they use desktop/clipboard resources.

## Fixtures and assertions

- COM behavior requires real Excel, not mocked `IExcelBatch`. Pure parsing,
  mapping, serialization, and generator tests need no Excel.
- Use a unique workbook per test. Do not combine `IClassFixture<T>` with a
  collection fixture on the same class: it can create competing Excel sessions.
- Follow nearby trait conventions: `Category`, `Feature`, `Layer`,
  `RequiresExcel`, `Speed`, and `RunType` where applicable.
- Assert workbook state and relevant returned fields, not only `Success`.
  Fixture COM access follows `excel-com-interop.instructions.md`.

## Save and round-trip behavior

Do not call `batch.Save()` for in-memory assertions. When testing persistence,
save/close and reopen in a new batch before asserting. Use `.xlsm` for VBA.

Test design and failure investigation: `tests/README.md`.
