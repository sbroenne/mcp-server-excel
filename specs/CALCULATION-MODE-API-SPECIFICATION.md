# Calculation Mode API Specification

## Source contract and scope

[ICalculationModeCommands and CalculationModeCommands](../src/ExcelMcp.Core/Commands/Calculation/CalculationModeCommands.cs)
define the actions, enums, results, and Excel implementation. Both entry points
route through the shared Service category `calculation`:

- MCP tool: `calculation_mode`, with `action` and `session_id`.
- CLI: `excelcli calculationmode <action> --session <sessionId>`.

Every action requires an open session. Mode is a property of that session's
Excel **application**, not an individual workbook. It affects all workbooks in
that Excel process. MCP and CLI sessions are separate.

## Actions and results

| Action | Additional parameters | Core result |
|--------|-----------------------|-------------|
| `get-mode` | None | `CalculationModeResult` |
| `set-mode` | Required `mode`: `automatic`, `manual`, or `semi-automatic` | `OperationResult` |
| `calculate` | Required `scope`: `workbook`, `sheet`, or `range`; conditional parameters below | `OperationResult` |

`get-mode` populates `mode`, `modeValue`, `calculationState`, `isPending`, and
the standard operation status/message fields. The result type also declares
`calculationStateValue`, `sheetName`, `rangeAddress`, and `scope`, but this
implementation does not populate them; do not treat them as observed state.

`set-mode` and `calculate` return operation status and a message, not
`previousMode`/`newMode` or a calculation-state payload. Successful results have
an empty or null `errorMessage`. Use `get-mode` to query mode separately.

| Mode | Excel value | Behavior |
|------|-------------|----------|
| `automatic` | -4105 | Automatic recalculation |
| `manual` | -4135 | Defer automatic recalculation; explicitly calculate before reading dependent results |
| `semi-automatic` | 2 | Automatic except data tables |

`calculate` has **no default scope** and can be used in any mode. It does not
change the mode.

| Scope | Required parameters | Excel call and extent |
|-------|---------------------|-----------------------|
| `workbook` | None beyond `scope` | `Application.Calculate()`: all open workbooks in the session's Excel application |
| `sheet` | MCP `sheet_name`; CLI `--sheet-name` | `Worksheet.Calculate()`: selected sheet in the session workbook |
| `range` | MCP `sheet_name`, `range_address`; CLI `--sheet-name`, `--range-address` | `Range.Calculate()`: selected range in the session workbook |

Missing sheet/range parameters produce a failed `OperationResult`. Invalid mode
or scope values are rejected. CLI parameter names are `--mode`, `--scope`,
`--sheet-name`, and `--range-address`.

## Bulk-write workflow

Use an existing session and worksheet; replace `SESSION_ID` with the identifier
returned by the same entry point. Check each operation's result before proceeding.

MCP calls:

```text
calculation_mode(action: "set-mode", session_id: "SESSION_ID", mode: "manual")
range(action: "set-values", session_id: "SESSION_ID", sheet_name: "Data", range_address: "A1:A2", values: [[10], [20]])
calculation_mode(action: "calculate", session_id: "SESSION_ID", scope: "workbook")
calculation_mode(action: "set-mode", session_id: "SESSION_ID", mode: "automatic")
```

Equivalent CLI commands:

```powershell
excelcli calculationmode set-mode --session SESSION_ID --mode manual
excelcli range set-values --session SESSION_ID --sheet-name Data --range-address A1:A2 --values '[[10],[20]]'
excelcli calculationmode calculate --session SESSION_ID --scope workbook
excelcli calculationmode set-mode --session SESSION_ID --mode automatic
```

Perform all intended writes between entering manual mode and calculating.
Explicitly return to automatic mode when finished; do not rely on session
closure to restore an earlier mode. If an operation fails while the session is
still usable, include that mode change in cleanup. A timed-out session must be
closed rather than reused. Calculation is not a save operation, and manual mode
does not guarantee timeout prevention or a particular performance improvement.

## References

- [Shared formula-calculation guidance](../skills/shared/gotchas.md#formulas-return-0-until-calculation-completes)
- [Calculation command tests](../tests/ExcelMcp.Core.Tests/Commands/Calculation/CalculationModeCommandsTests.cs)
- [Testing policy](../docs/ADR-001-NO-UNIT-TESTS.md)
- [Application.Calculation](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculation)
  and [XlCalculation](https://learn.microsoft.com/en-us/office/vba/api/excel.xlcalculation)
- [Application.Calculate](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculate),
  [Worksheet.Calculate](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.calculate),
  and [Range.Calculate](https://learn.microsoft.com/en-us/office/vba/api/excel.range.calculate)
- [Application.CalculationState](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculationstate)
