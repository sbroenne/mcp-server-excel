# Ranges and formatting

Values, formulas, and per-cell number formats use rectangular **2D arrays**.
Even a single value is `[[value]]`. Read before overwriting. Use content-only
clearing to preserve formats, and target only the changed cells.

Number display formats belong to range operations. Visual styles, validation,
merge/unmerge, widths, heights, and auto-fit belong to range-format operations.
For an existing `Sales` worksheet and a captured session:

```mcp
range(action: 'set-values', session_id: sessionId, sheet_name: 'Sales', range_address: 'A1:B2', values: [['Product','Amount'],['Widget',1250]])
range(action: 'set-number-format', session_id: sessionId, sheet_name: 'Sales', range_address: 'B2', format_code: '$#,##0.00')
range_format(action: 'format-range', session_id: sessionId, sheet_name: 'Sales', range_address: 'A1:B1', bold: true, fill_color: '#4472C4', font_color: '#FFFFFF')
range_format(action: 'auto-fit-columns', session_id: sessionId, sheet_name: 'Sales', range_address: 'A:B')
```

```cli
excelcli -q range set-values --session $sessionId --sheet Sales --range A1:B2 --values '[["Product","Amount"],["Widget",1250]]'
excelcli -q range set-number-format --session $sessionId --sheet Sales --range B2 --format-code '$#,##0.00'
excelcli -q rangeformat format-range --session $sessionId --sheet Sales --range A1:B1 --bold true --fill-color '#4472C4' --font-color '#FFFFFF'
excelcli -q rangeformat auto-fit-columns --session $sessionId --sheet Sales --range A:B
```

Check each result before continuing. The header example is for plain cells, not
an Excel Table. Use [Table styles](table.md) for Table headers/data. Combine all
visual properties in one call; use shared `format-ranges` for disjoint ranges on
one sheet. All target ranges are validated before that operation starts.
For new user-facing reports or requested formatting, see the scoped
[report-formatting workflow](report-formatting.md); preserve existing templates.

## Number formats and layout

| Meaning | US format code |
|---------|----------------|
| Number | `#,##0.00` |
| USD | `$#,##0.00` |
| Percentage | `0.00%` |
| Date | `yyyy-mm-dd` |
| Time | `hh:mm:ss` |
| Text | `@` |

Always supply US format codes. Excel translates separators and date codes for the
user's locale; different screenshot separators are not an error. Do not promise
a literal en-US rendering. After formatting, widen columns if values show
`#####`, while preserving intentional fixed layouts. Auto-fit rows for wrapped
text when needed.

`Good`, `Bad`, and `Neutral` are theme-aware styles with fills. Heading styles
provide hierarchy but no fill; use explicit visual formatting for colored
headers. `Normal` resets formatting.

## Formulas and merged cells

The server probes modern `Formula2` support once per session. Older Excel uses
`Formula`, with implicit intersection instead of dynamic-array spill behavior.
This does not add newer functions to Excel 2016/2019. Invalid formulas and
protected-cell errors fail rather than triggering a legacy retry.

Reads return canonical errors such as `#REF!` and `#DIV/0!`, with affected cells
and full formulas where available. Excel cannot reliably identify the precise
broken sub-reference; do not invent one.

Writes intersecting merged cells fail unless the target is just the merged
range's top-left cell. Write there for one merged value, or explicitly unmerge
before writing a grid.

## Links, comments, and names

Range-link actions manage external and internal hyperlinks. An internal target
uses a sub-address such as `'Summary'!A1`; removing a link preserves cell content.
On partial updates, omitted properties remain unchanged; an empty string clears
the URL, sub-address, or tooltip.

Threaded comments require a desktop Excel build exposing them. Local comment
text, author, dates, and replies are available; cloud mentions, assignments,
reactions, presence, sharing, and coauthoring are not.

Named ranges refer to cells, not literal values. Create the reference first,
then write its value. Listings omit hidden/internal names and avoid loading large
value previews; use a targeted read when values are required.
