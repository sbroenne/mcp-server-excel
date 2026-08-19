> **CLI syntax note:** This shared domain guide may use MCP-style `tool(action: ...)` examples as conceptual shorthand. Do not translate or paste those calls mechanically. Use the exact commands and kebab-case options in [cli-commands.md](./cli-commands.md) or live `--help`; notably, MCP `file` open/close maps to CLI `session` open/close, and MCP `worksheet` maps to CLI `sheet`.

# Dashboard & Report Best Practices

## The Professional Report Workflow

Every report or dashboard should follow this sequence:

```
1. Structure data → Excel Tables (never plain ranges)
2. Format values → Number formats by data type
3. Fit columns → auto-fit so nothing renders as #####
4. Add visuals → Charts with explicit positioning
5. Verify layout → Screenshot to confirm no overlaps
6. Save and close → Persist changes
```

## Step 1: Structure Data as Excel Tables

**Always use Excel Tables for tabular data:**

```
range(set-values, rangeAddress='A1', values=[[headers + data]])
table(create, tableName='SalesData', rangeAddress='A1:D20')
```

**Why Tables matter:**
- Auto-filters on every column
- Banded rows for readability
- Structured references in formulas
- Required for Data Model / DAX / PivotTables
- Auto-expand when new rows are added

## Step 2: Format Values by Data Type

**Apply number formats AFTER setting values — not before:**

| Data Type | Format Code | Result (en-US) |
|-----------|-------------|----------------|
| Currency (USD) | `$#,##0.00` | $1,234.56 |
| Currency (EUR) | `€#,##0.00` | €1,234.56 |
| Percentage | `0.0%` | 12.3% |
| Date | `yyyy-mm-dd` | 2025-01-22 |
| Number (thousands) | `#,##0` | 1,235 |
| Accounting | `_($* #,##0.00_)` | $ 1,234.56 |

**Always use US format codes** — Excel translates automatically to the user's locale.

**Rendered output is locale-dependent.** The `Result` column assumes en-US regional settings. Excel
interprets `,` and `.` in a format code according to the user's locale, so `$#,##0.00` displays as
`$1,234.56` on en-US but `$1.234,56` on de-DE. Same code, different separators — this is correct
behaviour, not a formatting bug. Never promise a literal rendering when reporting back to the user.

## Step 3: Fit Columns to Content

**Number formats make cells WIDER, so auto-fit immediately after formatting:**

```
range_format(action: 'auto-fit-columns', sheetName: 'Sales', rangeAddress: 'A:F')
```

A date formatted as `yyyy-mm-dd` or a currency value formatted as `$#,##0.00` does not fit the
default column width, so Excel renders the cell as `#####`. A screenshot taken before auto-fit will
show those columns as unreadable hash marks. Auto-fit before Step 5, not after.

Use `range_format auto-fit-rows` as well when any cell has `wrapText` enabled.

## Step 4: Position Charts with No Overlaps

**Charts have automatic collision detection and three positioning modes:**

### Single Chart (Auto-Position or targetRange)
```
# Option A: targetRange (explicit cell placement)
chart(create-from-range, sourceRange='A1:D20', targetRange='F2:K15')

# Option B: Omit position — auto-places below content
chart(create-from-range, sourceRange='A1:D20', chartType='Line')
# → Automatically positioned below the used range
```

### Multiple Charts (Dashboard) — Always Use targetRange
```
Place in a grid pattern below data:

Chart 1: targetRange='A22:F35'    (top-left)
Chart 2: targetRange='G22:L35'    (top-right)
Chart 3: targetRange='A37:F50'    (bottom-left)
Chart 4: targetRange='G37:L50'    (bottom-right)
```

### Collision Detection
All chart operations automatically warn about overlaps. If a result includes an `OVERLAP WARNING` message:
1. Use `chart(fit-to-range)` to reposition
2. Take `screenshot(capture, rangeAddress='A1:M50')` to verify

**Rules:**
- **Use targetRange for multi-chart layouts** — auto-positioning stacks vertically
- Leave 1-2 rows/columns gap between charts
- Place charts BELOW the data area, not beside it (more room)
- Keep chart sizes consistent (same row/column span)
- **Always check result messages** for overlap warnings

## Step 5: Verify with Screenshot

**Always take a screenshot after creating charts or complex layouts:**

```
screenshot(capture, rangeAddress='A1:M50')
→ Confirm: no overlaps, professional spacing, readable labels
→ Confirm: no column renders as ##### (if it does, go back to Step 3 and auto-fit)
→ If issues found: chart(fit-to-range) to reposition, then screenshot again
```

## Common Dashboard Layouts

### Summary Dashboard (Data + 2 Charts)
```
A1:D10    → Data table (formatted as Excel Table)
A12:F25   → Main chart (bar/column)
G12:L25   → Supporting chart (pie/line)
```

### Analytics Dashboard (4 Charts)
```
A1:D10    → Source data table
A12:F25   → Chart 1 (trend line)
G12:L25   → Chart 2 (distribution pie)
A27:F40   → Chart 3 (comparison bar)
G27:L40   → Chart 4 (detail scatter)
```

### Executive Report (Summary + Detail)
```
Sheet "Summary":
  A1:D5     → KPI table (small, formatted)
  A7:F20    → Summary chart

Sheet "Detail":
  A1:H100   → Full data table
  A102:H120 → Detail charts
```

## Formatting Checklist

- [ ] Data in Excel Tables (not plain ranges)
- [ ] Number formats applied (currency, dates, percentages)
- [ ] `range_format auto-fit-columns` run after formatting (no `#####` cells)
- [ ] Chart titles are descriptive
- [ ] Chart axis labels formatted (currency, percentages)
- [ ] No chart overlaps with data or other charts
- [ ] Consistent chart sizes in dashboards
- [ ] Screenshot taken to verify final layout
