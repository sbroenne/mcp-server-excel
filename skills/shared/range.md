# Ranges and formatting

**Mac experimental beta:** values/formulas, number formats, accepted copy/clear
variants, cell/row/column insertion and deletion, sizing, merge/unmerge, cell
locking, and named ranges are enabled. Row/column editing accepts single-area
ranges; disjoint row/column selections remain unavailable.
Rich styling, validation, links/comments, UsedRange/CurrentRegion, and merge-area
inspection remain unavailable. See
[macOS support](https://excelmcpserver.dev/macos-support/).

Values, formulas, and per-cell number formats use rectangular **2D arrays**.
Even a single value is `[[value]]`. Content writes reject occupied destinations
by default; use explicit permission for intentional replacement. Use content-only
clearing to preserve formats, and target only the changed cells.

Number display formats belong to range operations. Visual styles, validation,
merge/unmerge, widths, heights, and auto-fit belong to range-format operations.
For an existing `Sales` worksheet and a captured session:

```mcp
range(action: 'set-values', session_id: sessionId, sheet_name: 'Sales', range_address: 'A1:B2', values: [['Product','Amount'],['Widget',1250]])
range(action: 'set-number-format', session_id: sessionId, sheet_name: 'Sales', range_address: 'B2', format_code: '$#,##0.00')
range_format(action: 'auto-fit-columns', session_id: sessionId, sheet_name: 'Sales', range_address: 'A:B')
```

```cli
excelcli -q range set-values --session $sessionId --sheet Sales --range A1:B2 --values '[["Product","Amount"],["Widget",1250]]'
excelcli -q range set-number-format --session $sessionId --sheet Sales --range B2 --format-code '$#,##0.00'
excelcli -q rangeformat auto-fit-columns --session $sessionId --sheet Sales --range A:B
```

Check each result before continuing. For **Windows-only** header styling:

```mcp
range_format(action: 'format-range', session_id: sessionId, sheet_name: 'Sales', range_address: 'A1:B1', bold: true, fill_color: '#4472C4', font_color: '#FFFFFF')
```

```cli
excelcli -q rangeformat format-range --session $sessionId --sheet Sales --range A1:B1 --bold true --fill-color '#4472C4' --font-color '#FFFFFF'
```

The header example is for plain cells, not
an Excel Table. Use [Table styles](table.md) for Table headers/data on Windows. Combine all
visual properties in one call; use shared `format-ranges` for disjoint ranges on
one sheet. All target ranges are validated before that operation starts.
For new user-facing reports or requested formatting, see the scoped
[report-formatting workflow](report-formatting.md); preserve existing templates.

## Protected writes and copies

`set-values`, `set-formulas`, `copy`, `copy-values`, and `copy-formulas`
default to `reject-nonempty`. The server checks the direct destinations before
writing. This replaces a separate read used only to check whether cells are
empty; read when existing content is needed to understand the user's request.
For intentional replacement authorized by that request, use MCP
`overwrite_policy: 'allow'` or CLI `--overwrite-policy allow`.
Existing update scripts must add that option.

For an authorized update of the existing `Sales` worksheet:

```mcp
range(action: 'set-values', session_id: sessionId, sheet_name: 'Sales', range_address: 'B2', values: [[1500]], overwrite_policy: 'allow')
```

```cli
# Batch JSON uses "overwritePolicy": "allow".
excelcli -q range set-values --session $sessionId --sheet Sales --range B2 --values '[[1500]]' --overwrite-policy allow
```

Values, whitespace, zero, false, errors, and formulas displaying an empty string
are occupied. The same-value replacement is still an overwrite. Truly empty
cells pass even when formatted; blank incoming values or copied source blanks
do not authorize clearing existing content.

Conflicts return an error with the sheet and at most 10 cell addresses, noting
when further conflicts are not listed. No destination writes occur on rejection.
Failed inspection also stops the operation; never treat unknown content as
empty. Correct the destination or clarify unresolved intent. Do not automatically
retry with `allow`, clear the conflicting cells, or ask redundant confirmation
when replacement was already requested.

Copy checks expand a single-cell anchor to the source size. Larger destinations
must have row and column counts that are whole multiples of the source, and all
repeated-paste cells are checked. Protected copies require unmerged rectangular
source and destination ranges; ambiguous shapes and worksheet-edge overflow
fail before copying. `copy-formulas` follows Excel's formula-paste behavior:
source constants and blanks are copied too.

Value/formula payload dimensions must match a single rectangular destination.
Existing merged-write rules still apply. `allow` does not bypass Excel sheet
protection or other write errors. Formatting-only actions do not use this policy;
clearing and other tools' writes keep their own behavior.

Checking and writing execute in one session operation, but this is not a
transaction. Interactive users can still edit Excel. The check does not predict
future formula spills, dependent calculations, or table-generated changes
outside direct destinations, and a later Excel failure has no rollback promise.
Saving remains explicit.

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

Value/formula writes attempt to restore the prior calculation mode.
Restoration can fail without failing the write; use `get-mode` when subsequent
work depends on the mode. Automatic normally recalculates dependent formulas
after restoration; manual requires explicit calculation.
Semi-automatic excludes what-if data tables, not worksheet Tables. A successful
write does not guarantee completion of asynchronous refreshes or Python
calculations. Calculate and read back values when the result depends on them.

On Windows, the server probes modern `Formula2` support once per session. Older Excel uses
`Formula`, with implicit intersection instead of dynamic-array spill behavior.
This does not add newer functions to Excel 2016/2019. Invalid formulas and
protected-cell errors fail rather than triggering a legacy retry.

Reads return canonical errors such as `#REF!` and `#DIV/0!`, with affected cells
and full formulas where available. Excel cannot reliably identify the precise
broken sub-reference; do not invent one.

Writes intersecting merged cells fail unless the target is just the merged
range's top-left cell. Write there for one merged value, or explicitly unmerge
before writing a grid.

## Clearing ranges

`clear-all` removes values, formulas, and formats. `clear-contents` removes
values/formulas while preserving formats. `clear-formats` removes formatting
while preserving values/formulas. Each has no tool-level undo: check the exact
target before clearing.

These are in-memory changes until saved. An authorized close without saving can
discard them, but also discards any earlier unsaved work; it is not targeted undo.

## Finding matches

Find returns at most 10 matching cells by default. Set MCP `max_matches` or
CLI `--max-matches` to a positive whole number from 1 through 2147483647 to
change that limit. For the existing `Sales` worksheet and captured session:

```mcp
range_edit(action: 'find', session_id: sessionId, sheet_name: 'Sales', range_address: 'A1:B100', search_value: 'Widget', find_options: {}, max_matches: 5)
```

```cli
# Batch JSON uses maxMatches for the limit.
# --find-options and --replace-options accept JSON objects.
# For whole-cell matching, use --find-options '{"matchEntireCell":true}'.
excelcli -q rangeedit find --session $sessionId --sheet Sales --range A1:B100 --search-value Widget --find-options '{}' --max-matches 5
```

`matchingCells` contains the returned cell details. `totalCount` is the exact
number of matches, `returnedCount` is the number included, and `truncated` is
true only when matches were left out. No matches means an empty list, both
counts zero, and `truncated=false`. Exactly the limit is not truncated.

Excel still searches every match to count the total. The limit bounds retained
and returned cell details, not search time; requesting a large limit can
produce a large response. There is no paging or continuation.

## Links, comments, and names

Hyperlink and comment operations in this section are **Windows-only** in the
Mac beta. Named-range lifecycle is supported on both platforms.

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
