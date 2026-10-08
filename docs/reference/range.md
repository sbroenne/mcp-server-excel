# Ranges and formatting

Make targeted edits to the intended cells, preserving unrelated data and layout.
Use CLI help or MCP tool descriptions for current command syntax and inputs.
This guide explains choices that affect the workbook, not the command catalogue.

## Reusable cell styles

Use an existing style when it expresses the requested meaning. A custom cell
style can capture a visible source cell's formatting without changing that cell;
apply it separately to the intended destinations.

Changing a shared custom definition can affect every cell that uses it.
Deleting the style removes its name from existing users, with Excel deciding
what formatting remains. Built-in definitions are inspectable but not editable
through the custom-style lifecycle. Hidden source sheets are not made visible
as a workaround.

## Reusable table and dashboard styles

Table, PivotTable, slicer, and timeline styles share workbook definitions.
Clone a suitable definition when a reusable custom style is requested, then
apply it through the owning object's formatting controls.

These definitions support less formatting than ordinary cells. An unset
element is different from an element with explicit formatting, and Excel can
leave native properties unset after a change. Inspect the actual definition
before assuming that every requested property has been stored.

Custom updates can affect all existing users; deletion can remove their
formatting. Neither is a local change to just the selected report. Saving
remains explicit, and a native failure does not promise rollback.

## Cell protection

Locked cells and hidden formulas are enforced through worksheet protection.
Changing those cell flags alone does not protect the sheet. Formula hiding
affects Excel's UI, not tool inspection or file encryption; it is not secret
storage. See [worksheet protection](worksheet.md#protection-permissions).

## Row and column visibility

Visibility changes affect whole intersecting rows or columns, not just the
selected cells. Disjoint gaps remain unchanged. Hiding preserves stored
dimensions, although native size reads can report zero while hidden.

Excel cannot reliably distinguish manual hiding, filtering, zero size, and
collapsed groups. Filter and outline information provides context, not proof
of the cause. Showing rows does not clear filters or groups; those can hide
them again. Protection is not bypassed.

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

Content writes reject occupied direct destinations unless intentional replacement
is authorized. Values, whitespace, zero, false, errors, and formulas displaying
empty text are occupied. The same-value replacement is still an overwrite.
Formatting alone does not make a truly empty cell occupied.

After rejection, correct the destination or clarify unresolved intent. Do not
automatically grant overwrite permission, clear conflicting cells, or repeat
a confirmation when replacement was already requested. Failed inspection does
not establish that a destination is empty.

Copy scope can be larger than its starting address: a single-cell destination
expands to the source size, and larger destinations repeat the paste. Check the
full bounds and the effect of transposing. Copies require unmerged rectangular
areas with compatible dimensions.

Choose a content-only copy when formatting must remain unchanged. Formula
paste can also copy constants and blanks. Formats-only copying preserves
content but can transfer protection and conditional rules; validation-only
copying preserves content and visual formatting. Skipping empty source cells
does not skip formulas that merely display empty text.

Value and formula data must match the destination rectangle. Permission to
replace contents does not bypass sheet protection. Checks do not predict
future spills or Table-generated changes outside the direct destination.
Interactive edits remain possible, and a later Excel error can leave partial
changes. Copying uses Excel's clipboard; saving remains explicit.

## Number formats and layout

Display formatting is not data conversion. A percentage format does not turn
45 into 0.45, and a date format does not convert arbitrary text into dates.
Choose units and formats from the column's meaning, not its magnitude.
See [report formatting](report-formatting.md) for practical examples.

Use US number-format codes. Excel translates display separators for the user's
locale; different separators in a screenshot are not necessarily an error.
Read-back codes may change literal escaping while keeping the same meaning.
Explicit currency symbols remain explicit.

Widen columns when values show `#####`, unless a fixed template layout must
be preserved. Wrapped text may also need row auto-fit. Heading styles do not
necessarily include a colored fill; choose formatting that suits the request.

Inspect stored formatting when maintaining a template, and displayed formatting
when conditional rules affect appearance. A mixed-style range has no single
style: a summary fallback is not proof that every cell uses that style. Use
per-cell inspection when the distinction matters.

Validation replacement can remove the old rule before Excel rejects the new
one. Inspect the resulting rule after failure rather than assuming the previous
validation survived.

## Data validation

MCP `range_format` action `validate-range` and CLI `rangeformat validate-range`
replace the target's existing validation rule. Invalid validation types,
comparison operators, and error styles are rejected before removing that rule.
This does not promise rollback for errors Excel raises while applying a rule.

Explicit `show_input_message: false` / `--show-input-message false` and
`show_error_alert: false` / `--show-error-alert false` disable those messages.
Omitting these settings uses the documented defaults: input messages off,
error alerts on. Blank cells are allowed and list dropdowns are shown by default.
Use `get-validation` to inspect the resulting rule and message settings.

## Formulas and merged cells

Formula notation and cell-address notation are separate. Relative row/column
formulas are interpreted from each destination cell; changing their notation
does not change how target addresses are supplied.

Write a single merged value at the merged area's top-left cell. A grid write
intersecting merged cells can fail; unmerge only when that change is requested,
not as an automatic repair to a template.

## Ordinary-range and advanced filtering

Use the existing Table's filter when data is already a Table. Ordinary-range
filtering applies to an exact header-and-data rectangle and does not silently
replace a different filter elsewhere on the sheet.

Excel can normalize selected values into comparisons or combined criteria.
Read the actual criteria and visible rows rather than comparing the request's
original spelling with readback. Clearing ordinary criteria retains dropdowns.

Advanced filtering uses worksheet criteria and can filter in place or copy
results. Copied-output protection checks can cover more rows than ultimately
match. Excel does not expose the original advanced-filter criteria for faithful
readback. Clearing that state can affect worksheet-wide row filtering, so
authorize it explicitly rather than silently replacing it with an ordinary filter.

## Native data cleanup

Choose duplicate keys and header handling deliberately. Duplicate removal keeps
the first record for those keys, including blank records, and clears removed
rows within the selection rather than shifting cells below it.

For text-to-columns, protect identifiers, leading zeros, and dates from unwanted
conversion. Excel parses source formula text, not calculated formula results.
Check the full output area, including empty fields. Parsing in place can replace
the source; a separate output can overwrite neighboring data and needs permission
where occupied.

Preflight uses a short-lived, unsaved calculation workbook containing the
requested data, not a saved recovery copy. Native failures do not promise
rollback or bypass protection.

## Native fill and series

Copying a source edge and extending a pattern are different tasks. Pattern
extension includes the source and grows in one direction on the same sheet.
Formats-only extension preserves contents.

Trend fitting can change source values, so treat it as an intentional
replacement. Series safety checks can cover the whole possible destination
even if Excel stops early. Inspect the generated values rather than assuming
a successful call establishes the intended pattern.

## Native formula relationships

Native relationship inspection is not a complete workbook dependency graph.
It can omit cross-sheet and external references, miss dynamic references, or
report only an anchor rather than the effective target. An empty relationship
list does not prove that a cell has no dependencies.

Use the reported coverage and unresolved lookups. Inspection does not parse
formulas to invent missing links, recalculate, open external workbooks, select
cells, or create tracing arrows.

## Calculation and merged cells

Calculate and read back results when the task depends on newly written formulas.
Writes attempt to restore the previous calculation mode, but restoration can
fail without failing the write. See [calculation guidance](calculation.md).

Newer formula support depends on the Excel session. Using a legacy formula API
does not add modern functions to an older Excel installation, and failures are
not repaired by blindly retrying another notation.

Read cell-error details when formulas fail. Excel cannot always identify the
exact broken sub-reference. Spill inspection reflects current calculated state;
a blocked spill has no established intended output extent. Do not infer that
extent from an error or claim fresh results without calculation.

## Clearing ranges

Choose whether to remove contents, formats, or both. Clearing has no tool-level
undo. Discarding unsaved changes can reverse in-memory edits, but also loses
earlier unsaved work; it is not targeted undo.

## Finding matches

Choose text search for matching text and cell-category discovery for formulas,
errors, truly blank cells, or visible cells. Formulas displaying empty text
are not blanks, and visible-cell discovery excludes hidden rows and columns.

Search can return only part of its cell details while still reporting the exact
total. Check truncation before calling a list complete. Limiting returned
details does not limit the work required to count all matches.

## Links, comments, and names

Removing a hyperlink preserves its cell content. Internal links depend on real
sheet names and cell references.

Threaded comments require a supporting desktop Excel version. Local text and
replies are not Microsoft 365 mentions, assignments, reactions, sharing, or
coauthoring state.

Named ranges refer to cells. Define the reference before writing its value.
Discovery omits internal/hidden names and can omit large value previews; read
the intended name when values are needed. Changing a query parameter does not
itself refresh its load.

## Dates and reference spelling

Excel date serials are not Unix timestamps or Python ordinals. Check the
workbook's date system before external conversion, including the historical
1900 leap-year exception. Prefer display formatting when only readable dates
are needed.

Use actual worksheet names. In an Excel formula reference, a name with spaces
uses `'Sales Data'!A1`, not backticks.
