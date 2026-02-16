# Dashboard & Report Best Practices

## The Professional Report Workflow

Every report or dashboard should follow this sequence:

```
1. Structure data → Excel Tables (never plain ranges)
2. Format values → Number formats by data type
3. Add visuals → Charts with explicit positioning
4. Verify layout → Screenshot to confirm no overlaps
5. Save and close → Persist changes
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

| Data Type | Format Code | Result |
|-----------|-------------|--------|
| Currency (USD) | `$#,##0.00` | $1,234.56 |
| Currency (EUR) | `€#,##0.00` | €1,234.56 |
| Percentage | `0.0%` | 12.3% |
| Date | `yyyy-mm-dd` | 2025-01-22 |
| Number (thousands) | `#,##0` | 1,235 |
| Accounting | `_($* #,##0.00_)` | $ 1,234.56 |

**Always use US format codes** — Excel translates automatically to the user's locale.

## Step 3: Position Charts with No Overlaps

**CRITICAL: Every chart needs explicit positioning.**

### Single Chart
```
range(get-used-range) → "A1:D20"
chart(create-from-range, sourceRange='A1:D20', targetRange='F2:K15')
```

### Multiple Charts (Dashboard)
```
Place in a grid pattern below data:

Chart 1: targetRange='A22:F35'    (top-left)
Chart 2: targetRange='G22:L35'    (top-right)
Chart 3: targetRange='A37:F50'    (bottom-left)
Chart 4: targetRange='G37:L50'    (bottom-right)
```

**Rules:**
- **Never skip targetRange** — default positioning stacks charts on top of each other
- Leave 1-2 rows/columns gap between charts
- Place charts BELOW the data area, not beside it (more room)
- Keep chart sizes consistent (same row/column span)

## Step 4: Verify with Screenshot

**Always take a screenshot after creating charts or complex layouts:**

```
screenshot(capture-sheet)
→ Confirm: no overlaps, professional spacing, readable labels
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
- [ ] Column widths appropriate for content
- [ ] Chart titles are descriptive
- [ ] Chart axis labels formatted (currency, percentages)
- [ ] No chart overlaps with data or other charts
- [ ] Consistent chart sizes in dashboards
- [ ] Screenshot taken to verify final layout
