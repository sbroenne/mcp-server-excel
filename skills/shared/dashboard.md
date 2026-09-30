# Dashboard & Report Best Practices

## The Professional Report Workflow

For a requested new dashboard, adapt this sequence to the intended result.
Do not rebuild existing layouts or add visuals to a read-only task:

```
1. Structure data → Excel Tables when requested or needed
2. Format values → Number formats by data type
3. Fit columns → auto-fit so nothing renders as #####
4. Add visuals → Charts with explicit positioning
5. Verify layout → Check bounds/warnings; screenshot when the desktop is available
6. Save and close → When authorized; keep open if requested
```

## Step 1: Structure Data as Excel Tables

**Use Excel Tables when the source should be a structured dataset:**

```
range(set-values, range_address='A1', values=[[headers + data]])
table(create, table_name='SalesData', range_address='A1:D20')
```

**Why Tables matter:**
- Auto-filters on every column
- Banded rows for readability
- Structured references in formulas
- Required when adding worksheet data to the Data Model; ordinary PivotTables can also use ranges
- Auto-expand when new rows are added

## Step 2: Format Values by Data Type

**Choose number formats for the requested data; preserve intentional existing formats:**

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

**Check widths after formatting; auto-fit where it fits the intended layout:**

```
range_format(action: 'auto-fit-columns', sheet_name: 'Sales', range_address: 'A:F')
```

A formatted date or currency value may exceed the column width and display
`#####`. Adjust widths before checking the final layout; fixed-width report
columns may need a deliberate width rather than auto-fit.

Use `range_format auto-fit-rows` as well when any cell has `wrap_text` enabled.

## Step 4: Position Charts with No Overlaps

**Charts have automatic collision detection and three positioning modes:**

### Single Chart (Auto-Position or `target_range`)
```
# Option A: target_range (explicit cell placement)
chart(create-from-range, source_range_address='A1:D20', target_range='F2:K15')

# Option B: Omit position — auto-places below content
chart(create-from-range, source_range_address='A1:D20', chart_type='Line')
# → Automatically positioned below the used range
```

### Multiple Charts with a Grid Layout
```
Place in a grid pattern below data:

Chart 1: target_range='A22:F35'    (top-left)
Chart 2: target_range='G22:L35'    (top-right)
Chart 3: target_range='A37:F50'    (bottom-left)
Chart 4: target_range='G37:L50'    (bottom-right)
```

### Collision Detection
Chart create, move, and fit-to-range operations warn about overlaps. If a result includes an `OVERLAP WARNING` message:
1. Use `chart(fit-to-range)` to reposition
2. Check the revised bounds and, when available, capture the layout

**Rules:**
- Use `target_range` for a grid; auto-positioning is suitable for vertical stacks
- Leave 1-2 rows/columns gap between charts
- Place charts below or beside the data where there is sufficient room
- Keep chart sizes consistent (same row/column span)
- **Always check result messages** for overlap warnings

## Step 5: Verify with Screenshot

Use a screenshot for the requested visual result when an unlocked interactive
desktop is available. If unavailable, inspect bounds and overlap warnings and
report that visual verification was not performed; do not block an authorized
save/close solely because capture failed.

```
screenshot(capture, range_address='A1:M50')
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

- [ ] Data uses the requested structure (Tables when needed, otherwise ranges)
- [ ] Number formats applied (currency, dates, percentages)
- [ ] `range_format auto-fit-columns` run after formatting (no `#####` cells)
- [ ] Chart titles are descriptive
- [ ] Chart axis labels formatted (currency, percentages)
- [ ] No chart overlaps with data or other charts
- [ ] Consistent chart sizes in dashboards
- [ ] Final layout checked, or visual verification limitation reported
