# Report formatting

**Mac experimental beta:** number-format get/set and range sizing are enabled.
Rich cell formatting, Table/PivotTable/chart styles, and window freeze panes
remain gated, so the complete example below is Windows-only. Use the
[enabled action inventory](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/MACOS-ACTION-INVENTORY.md)
to select supported steps instead of treating the shared workflow as Mac parity.

Use this workflow when creating a new user-facing report or when formatting is
requested. It is not a requirement for reads, raw exports, or targeted data
updates. These are suggested defaults, not a universal financial or corporate
standard. They use existing operations; there is no `apply-report-formatting`
action.

## Establish scope before styling

The user's explicit instructions take precedence, followed by existing template
conventions, then these defaults. Inspect the intended sheets, data ranges,
number formats, and existing Tables before editing. Match an established
template; do not restyle unrelated cells or change a fixed layout.

Choose formats from the column's meaning and source metadata, not just its
values. A numeric column might be an identifier, a year, an amount, or a rate.
Ask only when unresolved currency, units, or value types would change the
result. Formatting changes display, not the meaning or precision of stored data.

## Use the owning object's formatting

- For an Excel Table, use [Table styles](table.md). Create a Table when requested
  or required, not merely because data is rectangular. Use preflight before
  creation when merged cells, headers, or sorting safety are uncertain.
- For plain cells, use [range formatting](range.md). Combine visual properties
  in one call, or use shared multi-range formatting for repeated styles.
- PivotTables and charts have their own formatting systems. Do not apply this
  plain-grid workflow to them; see [PivotTables](pivottable.md) and
  [charts](chart.md).

For a new report without an established style, use one consistent, readable
font, descriptive headers with units, and a restrained palette. Give plain
headers bold text and contrasting fill/text colours. Pair status colours with
labels; colour alone must not convey meaning.

Auto-fit the report's columns, not the entire worksheet. For long descriptions,
set a width suitable for the layout, enable wrapping, then auto-fit rows.
Preserve intentional widths in templates. Align numeric data right and text
left. Freeze the header when scrolling is useful; freeze counts include all
rows above the header, not just the header itself. See [worksheet views](window.md).

Avoid adding merged cells inside sortable data. Do not automatically unmerge an
existing template. Do not invent a "center across selection" alignment: the
range-format interface does not expose it. Meaningful sheet names help new
reports, but do not rename existing sheets or add cover sheets without need.

## Apply number formats without rewriting values

Use US format codes; Excel renders separators in the user's locale. Choose
precision appropriate to the data and preserve any requested currency or units.

| Column meaning | Suggested display | Important distinction |
|----------------|-------------------|-----------------------|
| Count | `#,##0` | Not for numeric identifiers or years |
| Decimal amount | `#,##0.00` | Display rounding must not round stored values |
| USD amount | `$#,##0.00` | Only for USD; state currency and scale in the header |
| Fractional rate | `0.0%` | `0.45` displays as 45%; formatting `45` displays as 4,500% |
| Native Excel date | `yyyy-mm-dd` | A display format does not convert arbitrary text into a date |
| Year | `0` | Avoid thousands separators; preserve text years if already text |
| Text or identifier | `@` | Preserve leading zeros and identifiers as text when writing |

Do not infer currency from magnitude or divide values by 100 merely because a
header contains `%`. A percentage-point value and a fractional rate are
different data types. Do not scale amounts to millions without an agreed
transformation and matching units. Applying `@` after a numeric write cannot
recover lost leading zeros or digits; supply identifiers as strings at creation.

### Plain-report example

The captured session contains a newly created `Report` sheet with plain cells
in `A1:C21`: description, amount in USD, and a rate stored as a fraction.
The user wants readable report formatting, and there is no template to preserve.
Check each result before the next step; a failed sequence is not rolled back.

```mcp
range_format(action: 'format-range', session_id: sessionId, sheet_name: 'Report', range_address: 'A1:C1', bold: true, fill_color: '#4472C4', font_color: '#FFFFFF')
range(action: 'set-number-format', session_id: sessionId, sheet_name: 'Report', range_address: 'B2:B21', format_code: '$#,##0.00')
range(action: 'set-number-format', session_id: sessionId, sheet_name: 'Report', range_address: 'C2:C21', format_code: '0.0%')
range_format(action: 'format-range', session_id: sessionId, sheet_name: 'Report', range_address: 'B2:C21', horizontal_alignment: 'right')
range_format(action: 'auto-fit-columns', session_id: sessionId, sheet_name: 'Report', range_address: 'A:C')
window(action: 'freeze-panes', session_id: sessionId, sheet_name: 'Report', frozen_rows: 1)
```

```cli
excelcli -q rangeformat format-range --session $sessionId --sheet Report --range A1:C1 --bold true --fill-color '#4472C4' --font-color '#FFFFFF'
excelcli -q range set-number-format --session $sessionId --sheet Report --range B2:B21 --format-code '$#,##0.00'
excelcli -q range set-number-format --session $sessionId --sheet Report --range C2:C21 --format-code '0.0%'
excelcli -q rangeformat format-range --session $sessionId --sheet Report --range B2:C21 --horizontal-alignment right
excelcli -q rangeformat auto-fit-columns --session $sessionId --sheet Report --range A:C
excelcli -q window freeze-panes --session $sessionId --sheet Report --frozen-rows 1
```

## Optional financial-model conventions

Use these only for a financial model where the user requests the convention or
the template already follows it, not for ordinary trackers or exports:

- Blue input text (`#0000FF`), black local formulas (`#000000`), green formulas
  linking other sheets (`#008000`), and red external-workbook links (`#FF0000`).
  External-link formulas take precedence over same-workbook links, which take
  precedence over purely local calculations. Identify formulas before colouring.
- Yellow fill (`#FFFF00`) can flag assumptions needing attention. Also label
  inputs, sources, and warnings so colour is not the only distinction.
- Parentheses for negatives and dashes for zero, for example
  `#,##0;(#,##0);"-"` or `0.0%;(0.0%);"-"`. Display multiples as `0.0"x"`.
- Keep scenario assumptions in separate cells and reference them in formulas.
  Record the source, date, and specific reference for sourced inputs in adjacent
  cells or supported comments. Never invent provenance or expose private source
  paths, credentials, or connection strings.

## Check the result before delivery

If the task creates or changes formulas, calculate in Excel and read back the
affected outputs with range value/formula reads. For a new formula-based report,
check every populated formula range, in manageable blocks if needed. Inspect
the returned cell-error details; syntax validation alone does not prove calculated
results are error-free. Restore the prior calculation mode after any temporary
change, including failure; see [working safely with Excel](behavioral-rules.md).

Fix errors introduced by the task. Report existing errors outside its scope
instead of silently rewriting unrelated formulas. Do not hide errors with
blanket `IFERROR` formulas or claim the entire workbook is error-free after
checking only a subset.

Check number formats and widths for truncated content or `#####`. Screenshots
are optional and require an interactive desktop; do not make them a gate for
unattended jobs. Save only the intended successful result using the existing
[session and saving rules](behavioral-rules.md). Report what was checked and any
remaining limitations.
