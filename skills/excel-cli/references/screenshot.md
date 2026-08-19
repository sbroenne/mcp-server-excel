> **CLI syntax note:** This shared domain guide may use MCP-style `tool(action: ...)` examples as conceptual shorthand. Do not translate or paste those calls mechanically. Use the exact commands and kebab-case options in [cli-commands.md](./cli-commands.md) or live `--help`; notably, MCP `file` open/close maps to CLI `session` open/close, and MCP `worksheet` maps to CLI `sheet`.

# Screenshot & Visual Verification Reference

## REQUIRED: Screenshot After Chart Creation

**You MUST call `screenshot` after creating any chart when visual output is requested or implied.** Do not close the file or end your response without capturing a screenshot.

```
1. chart(create-from-range, ...)  → Chart created
2. screenshot(capture, rangeAddress='A1:M20')  ← REQUIRED — never skip this step
3. file(close, save=true)
```

This rule applies even if:
- The chart was created on the first try
- No errors occurred
- The task description doesn't explicitly say "take a screenshot"

## Tools

- **`screenshot`**: Capture worksheet content as PNG images

## Actions

| Action | Purpose | Parameters |
|--------|---------|------------|
| `capture` | Capture a specific range | `rangeAddress` (default: A1:Z30), `sheetName`, `quality` |
| `capture-sheet` | Capture the worksheet's used cell range | `sheetName`, `quality` |

`capture-sheet` is cell-driven: it captures Excel's used range. On a chart-only worksheet, or when a chart extends beyond the used cells, use `capture` with an explicit range that covers the chart (for example, `A1:M25`).

## How Capture Works

Screenshots are taken from the live Excel window, so the image is exactly what Excel displays. Two consequences worth knowing:

- **An interactive desktop is required.** Excel is briefly shown and brought to the front. Capture fails on a locked desktop or a disconnected Remote Desktop session.
- **Protected sheets and hidden windows are fine.** Nothing is written to the workbook and the clipboard is not used, so capture never fails because a sheet is protected and never leaves a stray object behind.

Ranges larger than the Excel window are zoomed to fit and, if still too large, captured in several passes and stitched together. Very large ranges are truncated to their top-left portion; when that happens the result message says so — capture a smaller range for complete output.

## Quality Parameter

Default is `Medium` — use this for most cases. Only use `High` when fine text or formulas need careful inspection.

| Quality | Format | Scale | Size |
|---------|--------|-------|------|
| `Medium` | JPEG | 75% | ~4-8x smaller than High (default) |
| `Low` | JPEG | 50% | Smallest, good for layout overview |
| `High` | PNG | 100% | Full fidelity, largest file |

## When to Use Screenshots

### After Chart Creation or Positioning
```
1. chart(create-from-range, ..., targetRange='F2:K15')
2. screenshot(capture, rangeAddress='A1:O25')  → Verify chart doesn't overlap data
```

### After Complex Formatting
```
1. range(set-number-format, ...)
2. conditionalformat(add-rule, ...)
3. screenshot(capture-sheet)  → Verify formatting looks correct
```

### After PivotTable Layout Changes
```
1. pivottable(add-row-field, ...)
2. pivottable(add-value-field, ...)
3. screenshot(capture, rangeAddress='A1:M25')  → Include charts and verify layout
```

## Best Practices

1. **Verify chart placement**: After creating or repositioning charts, capture a screenshot to confirm no overlap with data or other charts
2. **Capture relevant area**: Use `capture` with a specific range rather than `capture-sheet` when you only need part of the worksheet
3. **Use after multi-step operations**: Screenshots are most valuable after a sequence of formatting, layout, or chart operations
4. **MCP returns image directly**: The image is returned as native ImageContent — no file handling needed
5. **Chart-only sheets need an explicit range**: `capture-sheet` uses the used cell range and may omit charts when no cells are used
6. **CLI with `--output`**: Use `excelcli screenshot capture --sheet <name> --range A1:M25 --output screenshot.png` to save the image directly
7. **Apply formatting once**: Apply each formatting operation (bold, fill color, number format) to a given range only once. Do not reapply unless a subsequent step explicitly changes or clears it — redundant calls waste turns and cost.

## Common Patterns

### Chart Overlap Verification
```
1. range(get-used-range) → "A1:D20"
2. chart(create-from-range, sourceRange='A1:D20', targetRange='F2:K15')
3. screenshot(capture, rangeAddress='A1:K20')
   → Visually confirm chart is positioned next to data, not on top of it
```

### Multi-Chart Dashboard Layout
```
When creating dashboards with multiple charts:

1. get-used-range → Know where data ends
2. Create Chart 1 with targetRange below/beside data
3. Create Chart 2 with targetRange that does NOT overlap Chart 1
4. Create Chart 3, Chart 4, etc. — each in a non-overlapping targetRange
5. screenshot(capture, rangeAddress='A1:M40') → Verify NO charts overlap each other or data

Key rules for multi-chart layouts:
- Use targetRange for every chart — never rely on default positioning
- Leave at least 1-2 rows/columns between charts
- Place charts in a grid pattern (e.g., 2x2) below the data area
- If overlap detected, use chart(fit-to-range) to reposition
```

### Dashboard Layout Check
```
1. Create multiple charts and tables
2. screenshot(capture, rangeAddress='A1:M40')
   → Verify overall dashboard layout, spacing, and alignment
3. If issues found: reposition with chart(fit-to-range), then screenshot again
```
