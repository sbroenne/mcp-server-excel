# Bug Report Test Plan

Source artifact: `excel-mcp-bug-report.md`

## Triage Summary

| Bug | Current classification | Current coverage | First test type to add |
|---|---|---|---|
| 1. Wide range write throws above 13 columns | Likely true regression | Partial Core regression only | Core regression, then MCP E2E |
| 2. New sheet rejects non-A1 writes until A1 is written | Likely true regression in worksheet-create workflow | No direct coverage on real workflow path | Core cross-command regression, then MCP E2E |
| 3. `format-range` lacks number format support | Not a regression; existing capability already exists via `range.set-number-format` | Strong Core coverage, no targeted MCP E2E | Acceptance only if team wants convenience on `format-range` |
| 4. No batch / multi-range formatting | Enhancement request, not regression | No coverage because feature does not exist | Acceptance tests after API decision |
| 5. No auto-fit column width support | Not a regression; capability already exists via `range_format.auto-fit-columns` | No explicit tests found | Positive MCP/Core coverage for existing feature, not regression |

## Bug 1

### Coverage today

- We already have one relevant Core regression: `RangeCommandsTests.Values.SetValues_WideHorizontalRange_NoOutOfMemoryError`.
- That test proves `SetValues` can write a single row across 16 columns (`A1:P1`) and read the values back.
- Coverage gap: it does not exercise the reported threshold (`14+` columns), it does not exercise multiple rows, and it does not prove the MCP tool path.

### First failing test to write

- Core integration regression: `SetValues_FourteenColumnsByFiveRows_WritesAndReadsBack`
- Arrange with a fresh sheet and a 5x14 payload targeting `A1:N5`.
- Assert actual round-trip values at the first cell, a middle cell, and the last cell. Do not stop at `Success = true`.

### Minimum coverage set

- Core regression: `A1:N5` round-trip succeeds.
- Core edge: `A1:M5` (13 columns) remains the control case so the threshold is explicit.
- Core edge: `B2:O6` succeeds to catch index math that depends on a non-A1 origin.
- MCP E2E: `range set-values` with the same 5x14 payload succeeds through the real transport and `get-values` returns all 14 columns.

### Regression or feature?

- Treat this as a real regression candidate.
- Existing functionality already promises arbitrary 2D range writes; the report describes a failure inside that promised behavior.

## Bug 2

### Coverage today

- We have indirect evidence that non-A1 writes can work on a fresh sheet: `RangeCommandsTests.Discovery.GetUsedRange_SheetWithSparseData_ReturnsNonEmptyCells` writes to `D10` on a newly added sheet.
- That does not cover the reported workflow, because the sheet is created by the test fixture helper, not by `SheetCommands.Create` or the MCP `worksheet create` action.
- MCP smoke coverage creates a sheet and then writes to `A1:C3`, which also misses the reported condition.

### First failing test to write

- Core cross-command regression: `CreateSheet_ThenSetValues_ToA3G10_WritesWithoutA1Priming`
- Flow: `SheetCommands.Create` -> `RangeCommands.SetValues` on `A3:G10` -> `RangeCommands.GetValues` on `A3:G10`.
- Assert the target range contains the exact payload and that `A1` remains empty, so the test proves no hidden priming write was required.

### Minimum coverage set

- Core regression: exact reported workflow on `A3:G10` with no prior write.
- Core edge: same workflow with `B2:D4` to show the bug is not tied only to row 3.
- Core control: write to `A1` after create still works, so the regression stays focused on non-A1 first writes.
- MCP E2E: `worksheet create` followed by `range set-values` to `A3:G10`, then `range get-values` verifies content.
- MCP E2E edge: repeat after reopening the session if sheet activation state looks session-sensitive.

### Regression or feature?

- Treat this as a real regression candidate until disproved.
- The system already supports worksheet creation and arbitrary range writes; requiring a hidden `A1` priming write would be a defect in the workflow contract, not a new feature.

## Bug 3

### Coverage today

- The report is inaccurate as stated. The product already supports number formatting via `range.set-number-format` and `range.set-number-formats`.
- Core coverage here is strong: `RangeCommandsTests.NumberFormat.*` plus `FormatTranslationTests.*` already verify currency, percentage, date, text, and display behavior.
- Gap: I did not find targeted MCP end-to-end tests proving that the tool-facing route is discoverable and works through the real protocol.

### First failing test to write

- No failing regression test should be written for the current report as-is.
- If the team wants `range_format.format-range` to accept a `number_format` convenience parameter, then the first failing test is an acceptance test: `RangeFormat_FormatRange_WithNumberFormat_AppliesDisplayFormat`.

### Minimum coverage set

- If no API change is planned: add one MCP E2E positive test for existing `range set-number-format` using a percentage display and verify the cell's displayed text or retrieved format.
- If an API change is planned: add acceptance tests for `format-range` with `number_format`, plus backward-compatibility tests proving existing `set-number-format` still works.

### Regression or feature?

- Not a regression.
- At most this is either a discoverability/documentation problem or a convenience-feature request to collapse two operations into one tool action.

## Bug 4

### Coverage today

- No existing coverage is expected because batch multi-range formatting does not exist in the current API.
- Current tests only cover single-range formatting behavior.

### First failing test to write

- Acceptance test, not regression: `FormatRangesBatch_MultipleNonContiguousRanges_AppliesExpectedFormats`
- This should be driven from the chosen public API shape, because the test contract changes depending on whether the team chooses `ranges`, `formats`, or named presets.

### Minimum coverage set

- MCP contract test for the chosen request shape.
- Core integration test applying identical formatting to multiple non-contiguous ranges in one call.
- Core integration test applying different formatting instructions per range if the API allows per-range overrides.
- Failure-behavior test that defines whether the batch is atomic or partially applied.
- MCP E2E test across a realistic report sheet with 3-5 formatting targets.

### Regression or feature?

- This is a feature request.
- The current API never promised batch formatting in one call, so there is no prior behavior to regress.

## Bug 5

### Coverage today

- The report is inaccurate as stated. The product already exposes `range_format.auto-fit-columns` and `range_format.auto-fit-rows`.
- I did not find explicit Core or MCP tests for auto-fit behavior.

### First failing test to write

- No failing regression test should be written for the current report as-is.
- The first useful test is a positive capability test: `AutoFitColumns_LongText_ExpandsColumnWidth`.
- If the product team wants an alias under `range` or `worksheet`, that becomes a separate acceptance test for the new alias rather than a regression.

### Minimum coverage set

- Core integration: write long text, capture width before and after `AutoFitColumns`, assert width increases.
- Core integration: `AutoFitRows` with wrapped text increases row height.
- MCP E2E: `range_format auto-fit-columns` works through the real tool path.
- MCP E2E edge: multi-column range such as `A:G` auto-fits all requested columns.

### Regression or feature?

- Not a regression.
- Current product capability exists already; the real issue is missing test coverage and probably discoverability in user guidance.

## Main concern with the bug report

This report mixes three different things:

1. Probable defects in existing promised behavior: Bugs 1 and 2.
2. Feature requests that need acceptance tests, not regressions: Bug 4, and Bug 3 only if the team chooses to add `number_format` to `format-range`.
3. Discoverability or reporting errors against existing API surface: Bugs 3 and 5 as written.

That mix matters because the wrong test strategy wastes time. Regressions need failing tests that lock the current contract. Enhancements need acceptance tests that define a new contract. Discoverability issues need positive end-to-end coverage plus docs and prompt guidance, not a fake regression narrative.