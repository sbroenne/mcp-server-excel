# Bugs 3, 4, and 5 Surface Investigation

Date: 2026-03-16
Requested by: Stefan Broenner
Agent: Cheritto

## Summary

- Bug 3 is not a missing backend capability. It is already supported under the `range` tool and reads as an API-shape/discoverability gap.
- Bug 4 is a real API gap. The surface does not currently support formatting multiple ranges in one action.
- Bug 5 is already supported under `range_format` and is primarily a discoverability/documentation gap.

## Bug 3: `format-range` Does Not Support Number/Percentage Format Strings

### Classification

API design gap / discoverability gap, not a missing feature.

### Current support status

Number formatting is already supported today under the `range` tool, not `range_format`.

- `src/ExcelMcp.Core/Commands/Range/IRangeCommands.cs`
  - `set-number-format`
  - `set-number-formats`
- `src/ExcelMcp.Core/Commands/Range/RangeCommands.NumberFormat.cs`
  - real COM implementation via `Range.NumberFormat`
- `specs/FORMATTING-VALIDATION-SPEC.md`
  - documents `set-number-format`, `set-number-formats`, and examples
- `skills/shared/range.md`
  - already tells agents to use `range` + `set-number-format` for dates, currency, and percentages
- `skills/shared/behavioral-rules.md`
  - already references `range set-number-format` in workflow guidance

The report is understandable because the surface is split in a way that is easy to miss:

- `range_format.format-range` handles visual styling
- `range.set-number-format` handles display formats

Top-level feature summaries currently make this split harder to discover:

- `FEATURES.md`
- `gh-pages/_includes/features.md`

Both currently reduce “Format data” to `range (format-range, validate-range)`, which hides number formatting entirely.

### Likely implementation path

Preferred path: fix discoverability first.

1. Update top-level feature/docs/skill guidance to make the split explicit:
   - visual styling -> `range_format`
   - number display formatting -> `range`
2. Only add a convenience alias if product wants a more intuitive API shape.

If a convenience alias is implemented, the smallest code path is to add an optional `numberFormat` parameter to `range_format.format-range` and internally delegate to the existing number format logic after visual formatting.

### Exact source/spec/docs likely affected if implemented

Code:

- `src/ExcelMcp.Core/Commands/Range/IRangeFormatCommands.cs`
- `src/ExcelMcp.Core/Commands/Range/RangeCommands.Formatting.cs`
- `tests/ExcelMcp.Core.Tests/Integration/Commands/Range/RangeCommandsTests.NumberFormat.cs`
- `tests/ExcelMcp.Core.Tests/Integration/Commands/Range/RangeCommandsTests.SetStyle.cs`

Parity / generated surfaces:

- MCP and CLI parity should flow automatically from the shared Core interface, but both generated surfaces must be validated after regeneration/build.
- No separate handwritten MCP-only or CLI-only feature branch should be introduced for this.

Docs/spec:

- `specs/FORMATTING-VALIDATION-SPEC.md`
- `skills/shared/range.md`
- `skills/shared/behavioral-rules.md`
- `skills/shared/anti-patterns.md`
- `FEATURES.md`
- `gh-pages/_includes/features.md`

### Scope recommendation

Keep this separate from Bug 4.

- PR 1A: docs-only discoverability fix
- PR 1B: optional alias/convenience API if docs-only fix is considered insufficient

## Bug 4: No Batch / Multi-Range Formatting Support

### Classification

Missing feature at the API surface.

### Current support status

There is no current action that applies formatting to multiple non-contiguous ranges in a single call.

What exists today:

- `range_format.format-range` can apply many properties in one call, but only to one range.
- `range_format.set-style` can apply one named style to one range.
- session batching keeps Excel open across many operations, which reduces workbook open/save overhead but does not reduce MCP/CLI action count.

That means the reported pain is real from an MCP/CLI round-trip perspective.

### Likely implementation path

Recommended first version:

- add a new `range_format` action that applies one shared formatting payload to many ranges
- shape: `ranges: []` plus the same formatting parameters already accepted by `format-range`

This is lower risk than a fully heterogeneous batch object list because it reuses the existing formatting contract and keeps validation simple.

Possible examples:

- `format-ranges` with shared properties + `ranges: ["A1:G1", "A3:G3"]`
- later follow-up: `format-ranges-batch` with `[{ rangeAddress, ...properties }]`

Generator feasibility looks acceptable. The existing surface already carries structured list parameters such as `IRangeEditCommands.Sort(... List<SortColumn> sortColumns ...)`, so a list-based request shape is consistent with current patterns.

### Exact source/spec/docs likely affected if implemented

Code:

- `src/ExcelMcp.Core/Commands/Range/IRangeFormatCommands.cs`
- `src/ExcelMcp.Core/Commands/Range/RangeCommands.Formatting.cs`
- likely a new helper/model type near the range command surface if a structured request object is used
- `tests/ExcelMcp.Core.Tests/Integration/Commands/Range/`
  - likely a new focused integration test file for batch formatting

Parity / generated surfaces:

- shared interface change will need regeneration/validation for both MCP and CLI
- operation-count and action-surface parity checks become mandatory under Rule 24

Docs/spec:

- `specs/FORMATTING-VALIDATION-SPEC.md`
- `skills/shared/range.md`
- `skills/shared/behavioral-rules.md`
- `skills/shared/anti-patterns.md`
- `FEATURES.md`
- `gh-pages/_includes/features.md`

If this ships as a net-new action, also expect documentation/count touchpoints that summarize operation totals.

### Scope recommendation

Keep this as its own PR.

- PR 2A: same-format multi-range action only
- PR 2B: heterogeneous per-range object payload only if needed after 2A

Do not bundle Bug 4 with Bug 3 or Bug 5 convenience aliases.

## Bug 5: No Auto-Fit Column Width Support

### Classification

Already supported under another action. Primary issue is discoverability.

### Current support status

Auto-fit support already exists on `range_format`:

- `src/ExcelMcp.Core/Commands/Range/IRangeFormatCommands.cs`
  - `auto-fit-columns`
  - `auto-fit-rows`
- `src/ExcelMcp.Core/Commands/Range/RangeCommands.Formatting.cs`
  - same partial-class family for formatting operations
- `specs/FORMATTING-VALIDATION-SPEC.md`
  - documents `auto-fit-columns` / `auto-fit-rows`

The bug report proposed putting auto-fit under `range` or `worksheet`, but the capability already exists in the more specific formatting tool.

The discoverability problem is credible because:

- the top-level feature summary hides auto-fit under a generic formatting bucket
- the shared skill guidance is much clearer about number formatting than it is about auto-fit

### Likely implementation path

Preferred path: documentation and skill surfacing only.

If product still wants the API shape from the report, treat it as an alias request rather than a new feature. The backend capability already exists.

### Exact source/spec/docs likely affected if implemented

Docs-first fix:

- `FEATURES.md`
- `gh-pages/_includes/features.md`
- `skills/shared/range.md`
- `skills/shared/behavioral-rules.md`

If an alias is added:

- `src/ExcelMcp.Core/Commands/Range/IRangeCommands.cs` or `src/ExcelMcp.Core/Commands/Range/IRangeFormatCommands.cs`
- `src/ExcelMcp.Core/Commands/Range/RangeCommands.Formatting.cs`
- integration tests for alias behavior
- `specs/FORMATTING-VALIDATION-SPEC.md`

### Scope recommendation

Keep Bug 5 out of the Bug 4 feature PR.

- PR 3A: docs/skills cleanup only
- PR 3B: optional alias if product explicitly wants a more discoverable command placement

## MCP / CLI Parity Notes

- The range and range_format surfaces are defined in shared Core interfaces and implementations.
- For these bugs, the main parity obligation is not separate handwritten MCP vs CLI logic; it is making sure both generated entry points expose the same action names, parameters, and docs after any interface change.
- Bug 3 and Bug 5 do not require new backend parity work unless aliases are added.
- Bug 4 would be a real new action and must be treated as a full parity change under Rule 24.

## Recommended review split

1. Docs/discoverability pass for Bugs 3 and 5
2. Optional convenience aliases as separate small PRs, only if still justified
3. Bug 4 as a standalone feature PR with integration tests and full doc/spec parity