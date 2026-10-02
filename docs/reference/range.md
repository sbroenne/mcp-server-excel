# Ranges and formatting

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
range_format(action: 'format', session_id: sessionId, sheet_name: 'Sales', range_addresses: ['A1:B1'], format_options: {bold: true, fillColor: '#4472C4', fontColor: '#FFFFFF'})
range_format(action: 'auto-fit-columns', session_id: sessionId, sheet_name: 'Sales', range_address: 'A:B')
```

```cli
excelcli -q range set-values --session $sessionId --sheet Sales --range A1:B2 --values '[["Product","Amount"],["Widget",1250]]'
excelcli -q range set-number-format --session $sessionId --sheet Sales --range B2 --format-code '$#,##0.00'
excelcli -q rangeformat format --session $sessionId --sheet Sales --range-addresses A1:B1 --format-options '{"bold":true,"fillColor":"#4472C4","fontColor":"#FFFFFF"}'
excelcli -q rangeformat auto-fit-columns --session $sessionId --sheet Sales --range A:B
```

Check each result before continuing. The header example is for plain cells, not
an Excel Table. Use [Table styles](table.md) for Table headers/data. Combine all
visual properties in one `format` call, including repeated styling across
multiple `range_addresses` (CLI `--range-addresses`). Supply the shared typed
payload as an MCP `format_options` object or CLI `--format-options` JSON string; its nested keys stay
camelCase. All target ranges and known invalid settings are checked before
writing; native failures do not promise rollback. The former `format-range`
and `format-ranges` actions and their scalar formatting inputs are removed.
Use theme indices/tints for colors that should follow `workbook.apply-theme`;
fixed RGB remains fixed. Each requested border is independent. Excel shares
cell boundaries, so an edge can also appear on the adjacent cell.
For new user-facing reports or requested formatting, see the scoped
[report-formatting workflow](report-formatting.md); preserve existing templates.

## Reusable cell styles

Use `workbook` actions `list-cell-styles` / `get-cell-style` to discover native
definitions. `create-cell-style` captures exactly one source cell, using MCP
`style_name`, `source_sheet_name`, and `source_cell_address` (CLI
`--style-name`, `--source-sheet-name`, `--source-cell-address`), without modifying
the source. The source worksheet must be visible; hidden sheets are rejected
without changing visibility. Its temporary activation restores the prior view.
Apply it separately with `range_format` / `rangeformat` `set-style`.

`update-cell-style` takes an MCP `style_options` object or CLI `--style-options`
JSON string. Nested keys stay camelCase; `formatOptions` reuses the shared
visual payload, except inside borders are range-only. Omitted inclusion flags
are preserved, including false. Changing a custom definition can affect every
existing user throughout the workbook. Built-in styles are read-only through
these lifecycle actions. `delete-cell-style` removes the custom name from
existing users; Excel determines retained formatting. No tool-level undo or
native-failure rollback is promised.

## Reusable table and dashboard styles

Use `workbook` `list-table-styles` / `get-table-style` to discover native
table/Pivot/slicer/timeline definitions, including all unset elements.
`create-table-style` clones an existing definition, using MCP `style_name` and
`source_style_name` (CLI `--style-name`, `--source-style-name`), without applying
it. `update-table-style` takes an MCP `table_style_options` object or CLI
`--table-style-options` JSON string. Nested keys remain camelCase.

Selected `elements` use exact native `elementType` names from inspection.
They support differential font emphasis/theme colors, solid fills, and six
outer/inside borders, but not font names/sizes, scripts or diagonal borders.
Only row/column stripe elements accept `stripeSize`; an unset stripe needs
formatting in the same request because its size alone does not create a native
definition. `clear: true` removes an
element and cannot be combined with formatting. Omitted elements/settings stay
unchanged. Native unset/default properties are reported rather than invented.

Apply a definition with `table` `set-style`, MCP `pivottable_calc` / CLI
`pivottablecalc` `set-layout-options`, or `slicer` `set-layout`. Custom updates
can affect all existing users. Built-in styles remain read-only.
`delete-table-style` can remove formatting from existing users; Excel determines
the fallback. Saving stays explicit; no native-failure rollback is promised.

## Cell protection

`range_link` (MCP) / `rangelink` (CLI) `set-cell-protection` changes only supplied
`locked` and MCP `formula_hidden` / CLI `--formula-hidden` flags. Named and
disjoint scopes preserve gaps. `get-cell-protection` returns both native flags
for every requested cell, with no first-cell or mixed-state fallback.
The old cell-lock actions are removed, not aliases.

```mcp
range_link(action: 'set-cell-protection', session_id: sessionId, sheet_name: 'Sales', range_address: 'B2:B10', locked: true, formula_hidden: true)
range_link(action: 'get-cell-protection', session_id: sessionId, sheet_name: 'Sales', range_address: 'B2:B10')
```

```cli
excelcli -q rangelink set-cell-protection --session $sessionId --sheet Sales --range B2:B10 --locked true --formula-hidden true
excelcli -q rangelink get-cell-protection --session $sessionId --sheet Sales --range B2:B10
```

Enforcement requires worksheet protection. Formula hiding affects Excel's UI,
not tool formula reads or file encryption. Do not describe it as secret storage.

## Row and column visibility

`range_format` `get-visibility` and `set-visibility` require MCP `axis` or CLI
`--axis` (`rows` or `columns`). Scope addresses select every intersecting whole
row or column, including named or disjoint ranges; gaps remain unchanged.
For `set-visibility`, MCP `hidden` or CLI `--hidden` must explicitly be true
or false. Hiding and showing preserve Excel's stored dimensions, but reads
report native current size, which can be zero while hidden. Row sizes are in
points; column sizes are in Excel character-width units.

Hidden cause is undetermined. Outline level, sheet filter mode, and ordinary
worksheet AutoFilter data-row membership are context, not proof of a manual,
filter, or grouping cause; table-filter membership is not exhaustive. Showing
does not clear filters or groups, and sheet protection is not bypassed.

```mcp
range_format(action: 'set-visibility', session_id: sessionId, sheet_name: 'Sales', range_address: 'A2:A4', axis: 'rows', hidden: true)
range_format(action: 'get-visibility', session_id: sessionId, sheet_name: 'Sales', range_address: 'A2:A4', axis: 'rows')
```

```cli
excelcli -q rangeformat set-visibility --session $sessionId --sheet Sales --range A2:A4 --axis rows --hidden true
excelcli -q rangeformat get-visibility --session $sessionId --sheet Sales --range A2:A4 --axis rows
```

## Protected writes and copies

`set-values`, `set-formulas`, and content-writing `copy` kinds
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

`copy` requires MCP `paste_kind` or CLI `--paste-kind`; batch JSON uses
`pasteKind`. Select `values` or `formulas` rather than the removed separate
copy actions. Native formula paste copies constants and blanks too. `formats`
preserves content but transfers number formats, protection, and applicable
conditional rules; `validation` preserves content and visual formatting.
Neither non-content kind requires overwrite permission.

Copy checks expand a single-cell anchor to the source's paste dimensions.
MCP `transpose` / CLI `--transpose` exchanges source rows and columns.
Larger destinations must be whole multiples, in both rows and columns, of those dimensions.
All kinds require unmerged rectangular source and destination ranges, under
either overwrite policy. Ambiguous shapes and worksheet-edge overflow fail
before copying. MCP `skip_blanks` / CLI `--skip-blanks` preserves destinations
corresponding to native empty source cells and excludes them from content
overwrite checks. Formulas displaying blank are not skipped. Both options
default to false. `destinationAddress` reports the complete paste bounds,
not an affected-cell count. Copy uses Excel's clipboard and clears owned copy
mode on exit; saving remains explicit.

```mcp
range(action: 'copy', session_id: sessionId, source_sheet: 'Sales', source_range: 'A1:B1', target_sheet: 'Sales', target_range: 'A10', paste_kind: 'formats')
```

```cli
excelcli -q range copy --session $sessionId --source-sheet Sales --source-range A1:B1 --target-sheet Sales --target-range A10 --paste-kind formats
```

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

For inspecting existing visual formatting, use MCP `range_format` `get-format`
or CLI `excelcli rangeformat get-format`. `view` / `--view` defaults to
`stored`; `displayed` includes conditional formatting; `both` returns both
snapshots. Every requested cell is included, without a preview limit or changing
selection. Each snapshot includes colors/themes, gradient stops, borders,
number format, alignment, protection, style, and dimensions. `mixedFields`
identifies mixed native properties rather than inventing a uniform value.
Use `get-style` when just the style name is sufficient.

```mcp
range_format(action: 'get-format', session_id: sessionId, sheet_name: 'Sales', range_address: 'A1:B100', view: 'both')
```

```cli
excelcli -q rangeformat get-format --session $sessionId --sheet Sales --range A1:B100 --view both
```

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

Number-format reads return Excel's canonical US codes; Excel may add or remove
literal escapes while preserving the display meaning. Reuse those returned codes
for subsequent writes. Currency literals stay explicit rather than changing to
the regional currency.

`Good`, `Bad`, and `Neutral` are theme-aware styles with fills. Heading styles
provide hierarchy but no fill; use explicit visual formatting for colored
headers. `Normal` resets formatting.

## Formulas and merged cells

`get-formulas` and `set-formulas` accept MCP `reference_style` / CLI
`--reference-style` (`a1` by default; `r1c1` for native row/column notation).
Batch JSON uses `referenceStyle`. This changes formula text, not range-address
notation. Relative R1C1 references are evaluated from each destination cell.
Excel's native modern or legacy formula API is selected from the session's
capabilities; failures never trigger a retry through another notation.

## Ordinary-range and advanced filtering

MCP `range_edit` / CLI `rangeedit` `apply-filter` selects an exact unmerged
header-plus-data rectangle and one relative MCP `column_index` / CLI
`--column-index`. Supply MCP `filter_options` / CLI `--filter-options`, not
the text parser's separate `options`. Nested keys remain camelCase.
The same native options work on Tables through `table_column` / `tablecolumn`;
do not apply ordinary-range filtering across a Table.

```mcp
range_edit(action: 'apply-filter', session_id: sessionId, sheet_name: 'Sales', range_address: 'A1:C100', column_index: 3, filter_options: {filterOperator: 'And', criteria1: '>=100', criteria2: '<=500'})
range_edit(action: 'get-filters', session_id: sessionId, sheet_name: 'Sales', range_address: 'A1:C100')
```

```cli
excelcli -q rangeedit apply-filter --session $sessionId --sheet Sales --range A1:C100 --column-index 3 --filter-options '{"filterOperator":"And","criteria1":">=100","criteria2":"<=500"}'
excelcli -q rangeedit get-filters --session $sessionId --sheet Sales --range A1:C100
```

Existing ordinary filters must match the exact rectangle; another filter is not
silently replaced. Reads preserve every column, arrays, native operators and
both criteria, including explicit native getter failures. Excel can normalize
short value selections into comparisons/OR. `clear-filters` clears the matching
ordinary criteria while retaining dropdowns.

`advanced-filter` uses native MCP `criteria_range` / CLI `--criteria-range`
with `mode: 'InPlace'` / `--mode InPlace` or `Copy`. Copy requires MCP
`copy_to_range` / CLI `--copy-to-range`; MCP `unique_only` / CLI `--unique-only`
requests native uniqueness. Copy output stays on the source worksheet.
Its overwrite preflight conservatively covers the full maximum source-row
extent, even when fewer records match. `checkedDestinationRange` is that
protection extent, not a claimed actual result extent.

Excel does not expose the original advanced row-filter criteria/scope.
`get-filters` reports worksheet row-filter context and
`advancedCriteriaAvailable: false`; it does not invent a faithful advanced
criteria read. Clearing that state requires explicit MCP `clear_advanced: true`
/ CLI `--clear-advanced true`, authorizing worksheet-wide native row-filter
clearing. No automatic retry, scope substitution, or protection bypass is allowed.

## Native data cleanup

MCP `range_edit` / CLI `excelcli rangeedit` `remove-duplicates` keeps the first
record for each explicit key. MCP `key_columns` / CLI `--key-columns` is a JSON
array of one-based indices relative to the selected rectangle. MCP `has_headers`
/ CLI `--has-headers` defaults to true and is never guessed. Counts exclude
headers but include blank records. Removed rows are cleared within the selection;
cells below it do not shift. Exact native counting requires one spare temporary
column, so 16383 distinct key columns is the supported maximum.

`text-to-columns` requires one source column, one same-sheet output anchor, and
typed `options`. MCP `source_range` / CLI `--source-range` selects input;
MCP `destination_cell` / CLI `--destination-cell` selects the anchor.
Nested option keys remain camelCase in both entry points. Delimited field
positions are one-based; fixed-width positions are zero-based and start at zero.
Native parsing, not a separate CSV parser, determines conversion and output.
Excel parses formula text, not calculated formula results.

The complete output, including empty fields, is checked before writing. Source
cells may be replaced when parsing in place; other occupied cells require MCP
`overwrite_policy: 'allow'` / CLI `--overwrite-policy allow`. Never automatically
retry with allow. Preflight uses a short-lived, unsaved native calculation
workbook containing only requested data; it is closed before the customer write
and the prior view is restored. No file copy or saved recovery workbook is needed.
Native failures do not promise rollback, bypass protection, or save edits.

```mcp
range_edit(action: 'remove-duplicates', session_id: sessionId, sheet_name: 'Sales', range_address: 'A1:C100', key_columns: [1,2], has_headers: true)
range_edit(action: 'text-to-columns', session_id: sessionId, sheet_name: 'Sales', source_range: 'E2:E100', destination_cell: 'G2', options: {comma: true, fields: [{position: 1, dataType: 'Text'}]})
```

```cli
excelcli -q rangeedit remove-duplicates --session $sessionId --sheet Sales --range A1:C100 --key-columns '[1,2]' --has-headers true
excelcli -q rangeedit text-to-columns --session $sessionId --sheet Sales --source-range E2:E100 --destination-cell G2 --options '{"comma":true,"fields":[{"position":1,"dataType":"Text"}]}'
```

## Native fill and series

Use MCP `range_edit` or CLI `excelcli rangeedit` for `fill`, `auto-fill`, and
`create-series`. `fill` copies contents and formatting from the selected
rectangle's source edge. `auto-fill` takes MCP `source_range` /
`destination_range` (CLI `--source-range` / `--destination-range`); the complete
destination includes the source and extends it in exactly one direction on the
same worksheet. `create-series` uses MCP `step_value` / `stop_value`
(CLI `--step-value` / `--stop-value`), with `orientation` selecting rows across
columns or columns down rows. Excel generates the pattern, not the server.

All three require unmerged rectangles. Content overwrite checks exclude the
source edge or AutoFill source. Formats-only AutoFill preserves contents.
Trend fitting can change source values and requires explicit `allow`.
For series with a stopping value, the safety check conservatively covers the
whole possible destination even when Excel stops early. These operations have
no rollback promise and do not bypass worksheet protection.

## Native formula relationships

MCP `range` and CLI `excelcli range` expose `trace-precedents` and
`trace-dependents`. Every starting cell in the exact requested scope and every
reachable native relationship is included, without a depth or preview cap.
`nodes` include current formulas/values; `edges` follow the requested traversal
direction; `cycles` identify strongly connected cells. Inspection does not
change selection/activation or recalculate.

This is not a complete workbook dependency graph. Native getters omit
cross-worksheet and external links, can omit `INDIRECT`, and can report only
`OFFSET`'s anchor instead of its effective target. Always read `coverage`:
`workbookComplete` is false even when `nativeTraversalComplete` is true.
Excel's no-range error cannot distinguish an empty relationship from unavailable
coverage; these lookups appear in `unresolved`. Do not interpret an empty
edge list as proof that a cell has no dependencies or dependents. The server
does not parse formula text, open external workbooks, or create tracing arrows.

## Calculation and merged cells

Value/formula writes attempt to restore the prior calculation mode.
Restoration can fail without failing the write; use `get-settings` when subsequent
work depends on the mode. Automatic normally recalculates dependent formulas
after restoration; manual requires explicit calculation.
Semi-automatic excludes what-if data tables, not worksheet Tables. A successful
write does not guarantee completion of asynchronous refreshes or Python
calculations. Calculate and read back values when the result depends on them.

The server probes modern `Formula2` support once per session. Older Excel uses
`Formula`, with implicit intersection instead of dynamic-array spill behavior.
This does not add newer functions to Excel 2016/2019. Invalid formulas and
protected-cell errors fail rather than triggering a legacy retry.

Reads return canonical errors such as `#REF!` and `#DIV/0!`, with affected cells
and full formulas where available. Excel cannot reliably identify the precise
broken sub-reference; do not invent one.

Use MCP `range` `get-spill-info` or CLI `excelcli range get-spill-info` for
native expanded-formula relationships, not a formula-text guess. Every
requested cell is included as `ordinary`, `source`, `result`, or `blocked`.
`sourceAddress` may be outside the requested scope. A blocked `#SPILL!` formula
has no established `spillAddress`, `spillRows`, or `spillColumns`; do not
invent an intended extent. Unsupported sessions fail explicitly. Inspection
does not recalculate or change selection and reflects the current calculated
state. Calculate separately when fresh results are needed.

```mcp
range(action: 'get-spill-info', session_id: sessionId, sheet_name: 'Sales', range_address: 'A1:B100')
```

```cli
excelcli -q range get-spill-info --session $sessionId --sheet Sales --range A1:B100
```

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

For a cell-category inventory, use MCP `range` `get-special-cells` with
`cell_kind`, or CLI `excelcli range get-special-cells` with `--cell-kind`.
Batch JSON uses `cellKind`. Unlike text `find`, discovery returns all matching
`areas` and their exact `cellCount` in the requested range, without a preview
limit or single-cell expansion. A named range uses an empty `sheet_name` /
`--sheet`; the result identifies its worksheet and resolved scope.

Formulas displaying empty text are not blank cells. Error discovery includes
stored errors and formula results. Visible discovery excludes hidden/filtered
rows and hidden columns. Empty `areas` with zero `cellCount` means no matches;
an error does not mean an empty range.

```mcp
range(action: 'get-special-cells', session_id: sessionId, sheet_name: 'Sales', range_address: 'A1:B100', cell_kind: 'errors')
```

```cli
excelcli -q range get-special-cells --session $sessionId --sheet Sales --range A1:B100 --cell-kind errors
```

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

## Dates and reference spelling

Date reads can return Excel serial numbers, not Unix timestamps or Python
ordinals. Check the 1900/1904 date system before external conversion; the 1900
calendar has a historical leap-year exception. Prefer display formatting when
only readable dates are needed.

Pass actual worksheet names in `sheet_name` (MCP) / `--sheet` (CLI), without
literal surrounding quotes. In a sheet-qualified Excel reference, spaces use
`'Sales Data'!A1`, not backticks.
