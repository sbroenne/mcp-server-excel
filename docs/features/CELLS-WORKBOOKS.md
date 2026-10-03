# Cells & Workbooks Features

Read, write, calculate, and format cells while managing worksheets, workbooks, named ranges, and files.

[← Back to the complete feature reference](../../FEATURES.md)

---

## 📁 File Operations (5 operations)

Open, create, and close Excel workbooks. Every other tool works on a session opened here.

**Operations:**
- **List Sessions:** View all active Excel sessions
- **Open:** Open workbook and create session (returns session ID for all subsequent operations). IRM/AIP-protected files are automatically detected and opened read-only with Excel visible for credential authentication — no extra parameters needed.
- **Close:** Close session with optional save
- **Create Empty:** Create new .xlsx or .xlsm workbook
- **Test:** Report existence, extension validity, openability, and IRM/AIP requirements through `canOpen`, `isIrmProtected`, `willOpenReadOnly`, and `requiresVisibleSession`. Ordinary workbooks are opened read-only in a temporary Excel session and closed without saving.

**Workflow:** List and match the intended workbook; reuse its session or
open/create; operate; list and check its `canClose`; close only when authorized
with an explicit save choice. MCP identifiers use `session_id` in inputs,
results, and list entries; CLI JSON uses `sessionId`.
Close without saving discards all unsaved edits, including earlier work, and
has no tool-level undo.

Operations within one session execute serially, but concurrent requests and
responses have no guaranteed order. Wait for dependent calls; different
sessions can run independently.

---

## 🧮 Calculation Mode (4 operations)

Control when and how Excel recalculates formulas — useful for speeding up bulk edits.

**Operations:**
- **Get Settings:** Read native mode/state, iteration settings, calculate-before-save, and the session workbook's precision-as-displayed flag
- **Set Settings:** Change only supplied application calculation settings and return native readback
- **Calculate:** Recalculate the owned application, sheet, or range; full and dependency-rebuild calculation require application scope
- **Set Precision:** Explicitly enable or disable workbook precision-as-displayed; enabling requires permission for permanent numeric precision loss

Mode and iteration settings affect all workbooks in the session's owned Excel
application, not other Excel processes. Application calculation also affects
all its open workbooks. The former `get-mode`, `set-mode`, and `workbook` scope
spellings are removed, not retained as aliases.
Enabling precision-as-displayed permanently rounds stored numbers to their
displayed precision across the workbook; disabling it does not restore lost digits.
Native failures do not promise rollback of supplied settings.

Value/formula writes attempt to restore the prior mode rather than always
forcing calculation. Restoration can fail without failing the write; use
`get-settings` when subsequent work depends on the mode. Automatic normally
recalculates dependent formulas after restoration; manual needs explicit
calculation. Semi-automatic excludes what-if data tables, not worksheet Tables.
Successful writes do not guarantee completion of
asynchronous refreshes or Python calculations. For bulk writes, remember and
restore the prior mode, including after failure.

---

## 📋 Ranges (64 operations)

Read and write cell values, formulas, and formatting across any range of cells.

**Formatting split:** use `range` for number display formats such as dates, currency, percentages, and text display. Use `range_format` for visual styling, validation, auto-fit, and size/layout changes. Its `format` action accepts one or more range addresses and one typed `formatOptions` payload, including number format when combined with visual changes.

**Fine formatting:** independent edge, inside, and diagonal borders; fixed RGB
or theme-index colors and tints; native underline kinds, strikethrough,
subscript/superscript and theme fonts; indentation, reading order, shrink-to-fit,
and center-across-selection. Omitted settings preserve native state. Targets,
protection, and known invalid options are checked before writing; native failures
do not promise rollback. Excel shares borders between neighboring cells.
The old `format-range` and `format-ranges` actions and scalar inputs are removed,
not kept as aliases. MCP uses `range_addresses` and `format_options`; CLI uses
`--range-addresses` and `--format-options`. Nested JSON keys stay camelCase.

**Data Operations:**
- **Get/Set Values:** Read or write cell values. Formula errors are returned as canonical names such as `#REF!`, with cell-level diagnostics preserving the formula and raw COM code.
- **Get/Set/Validate Formulas:** Read, write, or validate formula syntax across ranges. Formula reads include the same canonical error names and cell-level diagnostics as value reads.
- **Get Spill Info:** Read every requested cell's native expanded-formula source/result relationships and current result extent
- **Clear All/Contents/Formats:** Clear a range's contents, formats, or both
- **Copy:** Native all/values/formulas/formats/validation paste with transpose and skip-blank options
- **Insert/Delete Cells:** Shift cells to insert or remove space
- **Insert/Delete Rows:** Insert or delete entire rows
- **Insert/Delete Columns:** Insert or delete entire columns
- **Find:** Search a range for matching values, returning up to 10 cells by default with an exact total
- **Replace:** Find and replace values in a range
- **Sort:** Sort a range by one or more columns
- **Remove Duplicates:** Retain first records with explicit relative key columns and exact blank-aware counts
- **Text to Columns:** Native delimited/fixed-width parsing with complete destination protection, qualifiers, field types, and separators
- **Apply/Get/Clear Filters:** Typed native ordinary-range filters and complete native criterion reads
- **Advanced Filter:** Native criteria-range filtering in place or copied output, optionally unique
- **Fill:** Copy the source edge's contents and formatting down, up, left, or right
- **Auto Fill:** Extend native Excel patterns from an included source rectangle
- **Create Series:** Generate native linear, growth, date, or inferred series with step/stop options
- **Trace Precedents/Dependents:** Traverse all reachable native same-worksheet references with current formulas/values, cycles, and unresolved coverage

**Formula relationships:** MCP `range` and CLI `excelcli range` expose
`trace-precedents` / `trace-dependents`. All starting cells and reachable native
references are included without a depth/output cap. Results include `nodes`,
directed `edges`, strongly connected `cycles`, and `unresolved` lookups.
Native getters omit cross-worksheet/external references and do not fully resolve
dynamic references: `INDIRECT` can be omitted and `OFFSET` can report only its
anchor. `coverage.workbookComplete` is always false; `nativeTraversalComplete`
describes successful traversal of native getter results only.

A no-range native error can mean no relationships or unavailable relationships,
so it is reported in `unresolved`, not classified as empty. An empty edge list
does not establish absence of dependencies. Reads do not activate/select cells,
create tracing arrows, recalculate, parse formula text, or open external workbooks.

**Formula notation:** `get-formulas` and `set-formulas` accept MCP
`reference_style`, CLI `--reference-style`, or batch JSON `referenceStyle`.
Default `a1` preserves existing behavior; `r1c1` uses Excel's native relative
or absolute row/column formula notation. Range addresses remain A1. Relative
references are interpreted from each destination cell.

**Native cleanup:** MCP `range_edit` and CLI `excelcli rangeedit` expose
`remove-duplicates` and `text-to-columns`. Duplicate keys use MCP `key_columns`,
CLI `--key-columns`, or batch JSON `keyColumns`; indices are one-based relative
to the complete rectangle. Explicit header mode defaults to true. Counts exclude
the header, include blank records, and return retained bounds. Removed rows clear
inside the selection without shifting cells below it. Native exact counting
supports up to 16383 distinct key columns, leaving one temporary identity column.

Text parsing uses one source column and one same-sheet anchor, with typed
`options` whose nested keys remain camelCase. Delimited field positions are
one-based; fixed-width positions are zero-based and begin at zero. Excel parses
formula text rather than calculated formula results. Native preflight establishes
the full output, including empty fields, and closes its unsaved calculation
workbook before the requested write. Source cells may be replaced in place;
other occupied output cells require explicit `overwrite_policy: 'allow'` /
`--overwrite-policy allow`. Unknown options, invalid geometry, merged targets,
and worksheet overflow fail before customer writes. No file copy or saved
recovery workbook is required; protection and native failures are not bypassed.

**Native fills:** MCP `range_edit` and CLI `excelcli rangeedit` expose `fill`,
`auto-fill`, and `create-series`. All require unmerged rectangles. AutoFill's
destination must include its source and extend it in exactly one direction on
the same worksheet. Content writes protect occupied cells by default while
excluding source edges; formats-only AutoFill preserves contents. Trend fitting
can replace input values and requires `overwrite_policy: 'allow'` /
`--overwrite-policy allow`. Series use MCP `step_value` / `stop_value`, CLI
`--step-value` / `--stop-value`, or batch JSON `stepValue` / `stopValue`.
Stopping-value preflight conservatively protects the entire selected extent
even when Excel stops earlier. Native errors remain errors; no rollback or
protection bypass is promised.

**Find coverage:** MCP `range_edit(action: 'find')` accepts `max_matches`;
CLI `excelcli rangeedit find` accepts `--max-matches`; batch JSON uses
`maxMatches`. The default is 10, and any positive whole number through
2147483647 is accepted. Results retain `matchingCells` and include
`totalCount`, `returnedCount`, and `truncated`. A no-match result returns an empty
list, zero counts, and `truncated=false`; exactly the limit is not truncated.
Exact totals require searching every match even after the return limit is
reached. The limit bounds cell details, not search time. Paging is not provided.

**Protected writes:** `range` actions `set-values`, `set-formulas`, and content-writing
`copy` kinds default to `reject-nonempty`. Existing
values, whitespace, errors, and formulas displaying blank block the operation,
even for same-value replacements or blank incoming cells. Truly empty cells
pass regardless of formatting. For intentional updates, pass MCP
`overwrite_policy: 'allow'`, CLI `--overwrite-policy allow`, or batch JSON
`"overwritePolicy": "allow"`; existing update scripts must add this explicitly.
Neither policy bypasses Excel worksheet protection or existing write restrictions.

The check covers the whole direct destination before any write. Copies expand
single-cell anchors to the source size and check repeated-paste destinations.
All copy kinds require unmerged rectangles and destination dimensions that
are whole multiples of the source's paste dimensions, accounting for transpose;
ambiguous shapes fail before copying under either overwrite policy.
Excel formula paste also copies source constants and blanks.
Value/formula arrays must exactly match one rectangular target.

**Native paste:** `copy` requires MCP `paste_kind`, CLI `--paste-kind`, or
batch JSON `pasteKind`. Select the `values` or `formulas` kind for content-only
copies.
`formats` preserves values/formulas but transfers native number formats,
protection, and applicable conditional rules. `validation` preserves both
content and visual formatting. Neither non-content kind needs overwrite
permission. `transpose` / `--transpose` exchanges source rows and columns;
`skip_blanks` / `--skip-blanks` preserves destinations corresponding to
native empty source cells. Both default to false. Skipped blanks are excluded
from content overwrite checks; formulas displaying blank are not skipped.
Results identify resolved `sourceAddress` and expanded/repeated
`destinationAddress`, plus the selected options. Copy uses Excel's clipboard
and clears the owned application's copy mode on success or failure.

Conflicts report the sheet and at most 10 cell addresses, with an indication
when more exist, and make no destination writes. Inspection failures stop the
operation too. The check and write run together, but do not prevent interactive
Excel edits, predict future formula spills or table-generated changes outside
direct destinations, or promise rollback after a later Excel failure.
Do not automatically retry with `allow` after a rejection. Clearing, formatting,
and other tools' writes retain their own behavior; saving remains explicit.

**Clearing has no tool-level undo:** Clear All removes values, formulas, and
formats; Clear Contents preserves formats; Clear Formats preserves
values/formulas. Check the intended target before clearing.
Until saved, an authorized no-save close can discard changes, but also loses
earlier unsaved work.

**Discovery & Utilities:**
- **Get Used Range:** Get the worksheet's used range
- **Get Current Region:** Get the contiguous data region around a cell
- **Get Range Info:** Get a range's address and dimensions
- **Get Special Cells:** Find all formulas, constants, truly blank cells, errors, or visible cells in the requested scope

**Cell discovery:** MCP `range(action: 'get-special-cells')` takes `cell_kind`;
CLI `excelcli range get-special-cells` takes `--cell-kind`; batch JSON uses
`cellKind`. Results contain the resolved `sheetName`, absolute `rangeAddress`,
`cellKind`, all matching rectangular `areas`, and the exact `cellCount`.
There is no preview limit. Separate or overlapping input areas use their actual
union, without including gaps or counting a cell twice.

A single-cell request stays one cell. Blank discovery includes requested cells
outside the used range; formulas displaying empty text are not blanks.
Errors include stored error values and formula results. Visible discovery
excludes hidden rows and columns, including filtered rows. No matches returns
empty `areas` and zero `cellCount`; invalid inputs and Excel failures remain
errors. Discovery does not select cells or change workbook content.

**Expanded-formula inspection:** MCP `range(action: 'get-spill-info')` and CLI
`excelcli range get-spill-info` return every requested cell as `ordinary`,
`source`, `result`, or `blocked`, with native `sourceAddress`, `sourceFormula`,
`spillAddress`, `spillRows`, and `spillColumns` where established. A result
cell can identify a source outside the requested scope. A formula returning
`#SPILL!` identifies itself but has no invented intended extent. `capability`
is `supported` on success; unsupported Excel sessions return an explicit error,
not an empty successful inspection. Reads do not recalculate or change
selection, so they reflect Excel's current calculated state. There is no
preview limit. Named ranges and exact unions of disjoint scopes are supported.

**Hyperlinks:**
- **Add Hyperlink:** Add a hyperlink to a cell
- **Update Hyperlink:** Change an external/internal target, display text, or tooltip
- **Remove Hyperlink:** Remove a hyperlink
- **List Hyperlinks:** List all hyperlinks in a range
- **Get Hyperlink:** Get a specific hyperlink's target

**Threaded Comments:**
- **Add Threaded Comment:** Add a top-level modern comment to one cell
- **List Threaded Comments:** Read a cell's comment and replies
- **Add Threaded Comment Reply:** Reply to an existing cell comment
- **Delete Threaded Comment:** Delete a comment thread and its replies

**Number Formatting (`range`):**
- **Get Number Formats:** Read number formats as a 2D array
- **Set Number Format:** Apply one number format uniformly
- **Set Number Formats:** Apply individual per-cell number formats

**Visual Formatting (`range_format`):**
- **Get Format:** Read complete stored, displayed, or both formatting snapshots for every requested cell
- **Get Style:** Read the applied cell style
- **Set Style:** Apply a built-in Excel style
- **Format Range:** Set font, color, borders, alignment, orientation
- **Format Ranges:** Apply one shared formatting payload to multiple ranges

**Formatting inspection:** MCP `range_format(action: 'get-format')` and CLI
`excelcli rangeformat get-format` accept `view`: `stored` (default),
`displayed`, or `both`. Displayed formatting includes conditional rules without
changing stored formatting. Results include all requested cells with their
addresses, fonts, colors/themes/tints, fills and complete gradient stops,
eight border positions, invariant number formats, alignment, wrapping,
indentation, protection, style, and dimensions. Native mixed properties, such
as differing characters within a cell, are listed in `mixedFields` rather than
replaced with made-up defaults. There is no preview limit or selection change.
Named ranges use an empty sheet name; disjoint and overlapping scopes return
their exact union. Use `get-style` if only the style name is needed.

**Data Validation (`range_format`):**
- **Add Validation:** Add dropdown, number/date/text validation rules
- **Get Validation:** Read current validation info
- **Remove Validation:** Remove validation rules

**Merge Operations (`range_format`):**
- **Merge Cells:** Merge a range into one cell
- **Unmerge Cells:** Undo a merge
- **Get Merge Info:** Read current merge state

**Cell Protection:**
- **Set Cell Protection:** Change only supplied lock/formula-hiding flags in exact named or disjoint scopes
- **Get Cell Protection:** Read both flags for every requested cell, without a first-cell or mixed-state fallback

These replace the former cell-lock actions. Enforcement requires sheet protection;
formula hiding affects Excel's UI, not tool inspection or file encryption.

**Sizing & Auto-Fitting (`range_format`):**
- **Auto-Fit Columns / Rows:** Resize columns or rows to fit content
- **Set Column Width / Row Height:** Set column widths in Excel character-width units and row heights in points
- **Get Visibility:** Read every unique whole row or column intersecting an exact scope, including native current size, outline level, and ordinary worksheet AutoFilter data-row membership
- **Set Visibility:** Hide or show intersecting whole rows or columns, preserving stored dimensions and leaving disjoint gaps unchanged

Hidden cause is reported as undetermined: Excel COM does not reliably distinguish
manual hiding, filtering, zero size, and collapsed groups. Filter and outline facts
provide context, not proof of the cause; table-filter membership is not exhaustive.
Hidden dimensions can report native current size zero. Showing dimensions does
not clear filters or groups; those can affect visibility again. Sheet protection
is not bypassed, and native failures do not promise rollback.

---

## 📄 Worksheets (35 operations)

Add, rename, move, and manage worksheets — including tab colors, visibility, protection, legacy cell notes, inline images, shapes, and page setup.

**Lifecycle:**
- **List:** List worksheets in the workbook
- **Create:** Add a new worksheet
- **Rename:** Rename a worksheet
- **Copy:** Copy a worksheet within the workbook
- **Move:** Move a worksheet within the workbook
- **Delete:** Remove a worksheet

**Cross-Workbook Operations:**
- **Copy to File:** Copy a worksheet to another workbook (atomic)
- **Move to File:** Move a worksheet to another workbook (atomic)

**Delete has no tool-level undo:** It removes all sheet contents and can break
dependent references. Check the intended sheet and its dependencies.
**Move to File has no tool-level undo:** It removes the source sheet and saves
both files. Closing another session without saving cannot
reverse the saved transfer.

**Tab Colors:**
- **Set Tab Color:** Set a worksheet tab's RGB color
- **Get Tab Color:** Read the current tab color
- **Clear Tab Color:** Reset the tab to its default color

**Visibility:**
- **Show:** Make a worksheet visible
- **Hide:** Hide a worksheet (still shown in the Unhide dialog)
- **Very Hide:** Hide a worksheet from the Excel UI entirely
- **Get Visibility:** Read the current visibility status
- **Set Visibility:** Set visibility status directly

**Protection:**
- **Set Protection:** Protect/unprotect with native component, permission, and selection settings; supplied options replace the configuration
- **Get Protection:** Read protected components, all native permissions, selection restrictions, and runtime-only UI protection

Unspecified permissions use Excel's restrictive defaults. Filtering permission
allows changing existing filters; sorting/deleting still requires unlocked cells.
UI-only protection permits automation but is not persisted after reopening.
Passwords are not returned. Native failures do not promise rollback.

**Cell Notes:**
- **Set Comment:** Create or update a legacy cell note through Excel's Comment COM API
- **Get Comment:** Read the current legacy cell note text
- **Clear Comment:** Remove a legacy cell note

**Images:**
- **Add Image:** Insert an image from disk and anchor it to a cell
- **Get Image Count:** Read how many images are currently on a worksheet

**Shapes:**
- **Add Shape:** Insert a basic rectangle shape and anchor it to a cell
- **Get Shape Count:** Read how many shapes are currently on a worksheet

**Page Setup:**
- **Set Page Setup:** Configure orientation, centering and fit modes, plus typed print areas, repeated title rows/columns, point margins, headers/footers, paper, page order, gridlines/headings, fixed zoom and native printing options. Omitted settings remain unchanged; empty text explicitly clears. Fixed zoom and fit parameters cannot be combined; zero fit counts mean unlimited.
- **Get Page Setup:** Read native print scope, margins, headers/footers, paper, order, active scaling and printing flags. Fixed zoom and fit values are nullable when inactive.
- **Get Page Breaks:** Read every horizontal/vertical break exposed by Excel in the current print scope, including position, manual/automatic state and extent. Printer and scaling affect automatic breaks.
- **Set Page Breaks:** Replace all manual worksheet breaks using required row/column lists; empty lists clear. Validate distinct positions before resetting; automatic breaks remain. This also replaces manual breaks outside the current print scope.

Page setup does not print or open a modal preview. Existing PDF/XPS export honors
print areas and settings by default. Margins use points, titles require complete
rows/columns, and paper support depends on the installed printer driver.

**Outlines:**
- **Group / Ungroup:** Group complete row or column ranges and remove one grouping level
- **Get Outline Info:** Read outline level, hidden state, summary positions, and automatic styles
- **Set Outline Settings:** Configure summary rows/columns and automatic styles
- **Show Outline Levels:** Expand or collapse row and column groups to requested levels
- **Clear Outline:** Remove all row and column groups

---

## 📘 Workbook (27 operations)

Manage workbook metadata, protection, document properties, file variants, exports, and external links.

**Operations:**
- **Set Protection:** Protect or unprotect the current workbook, optionally with a password
- **Get Protection:** Determine whether the current workbook is protected
- **Set View Options:** Toggle workbook window gridlines and headings on or off
- **Get Theme:** Read all 12 native theme colors and major/minor Latin, East Asian, and complex-script font definitions; empty script fonts remain empty, and unavailable theme/scheme names are explicit.
- **Apply Theme:** Apply an existing absolute-path Office `.thmx` file through Excel. Theme-sensitive formatting changes throughout the workbook; fixed RGB colors remain fixed. Saving remains explicit.
- **List Cell Styles:** Every native name, localized name, and built-in/custom status; no cap
- **Get Cell Style:** Complete native definition, six supported borders, inclusion flags, and explicit read limitations
- **Create Cell Style:** Capture one visible source cell's stored formatting as a new custom style without modifying the source; restore the previous view after temporary worksheet activation
- **Update Cell Style:** Change supplied custom-style settings and preserve omitted inclusion flags; existing users throughout the workbook can change
- **Delete Cell Style:** Remove a custom style; existing users lose its name and Excel determines retained formatting
- **List Table Styles:** Every native table/Pivot/slicer/timeline style name, built-in/custom status, and availability flag; no cap
- **Get Table Style:** All 43 distinct native elements, including unformatted elements, differential font/fill/border settings, applicable stripe sizes, and explicit unset/unavailable fields
- **Create Table Style:** Clone an existing native style into a new custom definition without applying it
- **Update Table Style:** Change selected native elements, clear an element, set row/column stripe sizes, and change availability; omitted settings and elements are preserved
- **Delete Table Style:** Remove a custom style; existing tables, PivotTables, slicers, and timelines can lose its formatting, with Excel determining the fallback
- **Get View Options:** Read back workbook window gridlines and headings state
- **Get Info:** Read workbook name, path, format, saved/read-only state, and protection metadata
- **List/Get/Set/Delete Document Properties:** Manage built-in and custom workbook properties
- **Save As:** Save as `.xlsx`, `.xlsm`, `.xlsb`, or `.xls` and move the active session to the new path
- **Save Copy As:** Create a same-format copy without changing the active workbook
- **Export Fixed Format:** Publish PDF or XPS with quality, page-range, and print-area controls
- **List/Update/Break External Links:** Inspect, refresh, or permanently replace linked-workbook formulas

Built-in cell and table styles are inspectable but read-only through style
lifecycle operations. Apply a cell style separately with `range_format`
`set-style`. Cell styles have four outer borders and two diagonals; inside borders
are range-only. Column widths, row heights, and unavailable native
`Style.MergeCells` are not invented in style reads. Hidden source worksheets are
rejected without changing their visibility.

Table-style elements support font emphasis, theme fonts/colors/tints, solid
fills, and four outer/two inside borders, but not font names/sizes, subscript,
superscript, alignment, number formats, or diagonal borders. Unformatted elements
have explicit null formatting and stripe size. Set formatting with `stripeSize`
when defining an unset stripe; a size alone does not create its native format.
Native differential properties
can remain unset after a false/default write; inspection reports that state
rather than inventing a value. Excel can slightly normalize tint values.
Apply a custom definition with `table` `set-style`, `pivottable_calc`
`set-layout-options`, or `slicer` `set-layout`.

Custom-style updates can affect existing users throughout the workbook. Native
failures do not promise rollback, and deletion has no tool-level undo.

**Break External Link has no tool-level undo:** Linked formulas become their
current values.

> Printing and print preview are intentionally excluded because physical printer output and modal preview are unsafe for unattended automation.

---

## 🏷️ Named Ranges (Parameters) (6 operations)

Manage named ranges — ideal for driving workbook parameters that Power Query and formulas react to.

**Operations:**
- **List:** List visible user-defined named ranges with references; hidden/internal Excel names (including Power Query `ExternalData_*` and AutoFilter names) are omitted before value inspection, and large ranges return metadata without materializing values
- **Read:** Get value of a named range
- **Write:** Set a named-range value. Invariant numeric and Boolean strings become typed Excel values; identifiers such as `2.0.13` remain text.
- **Create:** Create new named range
- **Update:** Modify existing named range
- **Delete:** Remove named range

**Notes:**
- **Use cases:** Manage workbook parameters that formulas and Power Query can read. Updating a parameter does not itself refresh Power Query; explicitly refresh the affected query when its results must reflect the new value.

---

## Related feature areas

- [Data & analytics](DATA-ANALYTICS.md) — transform ranges and tables with Power Query, DAX, and PivotTables
- [Charts & visualization](CHARTS-VISUALS.md) — turn workbook data into charts and visual reports
- [Automation & advanced](AUTOMATION-ADVANCED.md) — automate workbooks with VBA, Python, and What-If Analysis
- [Example workflows](../USE-CASES.md) — see these capabilities combined in practical requests
- [Installation](../INSTALLATION.md) — choose and configure the MCP Server or CLI

## Task guides

- [Real Excel automation vs. file-parser libraries](../guides/EXCEL-COM-VS-FILE-PARSERS.md)
- [Refresh Power Query from an AI assistant](../guides/REFRESH-POWER-QUERY.md)
