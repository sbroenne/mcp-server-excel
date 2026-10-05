# Cells & Workbooks Features

Update live workbook data and formulas, format reports, and manage the sheets,
styles, and files around them.

[Back to the feature overview](../../FEATURES.md)

These are capability summaries. Use CLI help or MCP tool descriptions for
current command details, and the linked guidance for workflow decisions.

---

## File Operations (5 operations)

- **Start work:** Open an existing workbook or create a new one and retain its Excel session.
- **Discover sessions:** Find the intended open workbook and reuse its session.
- **Check access:** Test whether Excel can open a file and identify protection or authentication requirements.
- **Finish deliberately:** Save authorized changes or close without saving.

CLI and MCP sessions are separate. Reuse the matching session, wait for
dependent operations, and close only after active work finishes. Discarding
changes loses all unsaved edits, including earlier work; it is not targeted undo.
IRM/AIP-protected files need visible Excel authentication. Excel determines the
signed-in user's editing rights; protection alone does not force read-only access.

[Session and saving guidance](../reference/behavioral-rules.md#sessions-and-failures)

---

## Calculation Mode (4 operations)

- **Inspect calculation:** Read Excel's calculation state and settings.
- **Control recalculation:** Adjust calculation and iteration settings or recalculate the application, a sheet, or a range.
- **Manage stored precision:** Explicitly change the workbook's precision-as-displayed setting when that permanent change is intended.

Calculation settings affect every workbook in the session's Excel application.
For bulk edits, restore the previous calculation mode even after failure.
A successful write does not establish that asynchronous refreshes or Python
calculations have finished.

Precision-as-displayed changes stored numbers, not merely their appearance.
Disabling it cannot restore lost digits.

[Calculation workflow guidance](../reference/calculation.md)

---

## Ranges (64 operations)

- **Work with data:** Read and write values or formulas, search and replace, copy selected content, and insert or remove cells, rows, and columns.
- **Discover workbook state:** Inspect used regions, matching cell categories, formula errors, spilled arrays, and native formula relationships.
- **Clean and organize:** Sort, filter, remove duplicates, split text into columns, and extend data with Excel's fills and series.
- **Format cells:** Apply number formats, cell styles, fonts, fills, borders, alignment, and shared formatting across multiple ranges.
- **Control layout and input:** Set sizes and visibility, merge cells, and manage validation and cell-protection settings.
- **Add context:** Manage hyperlinks and supported threaded comments.

Content writes protect occupied destinations unless intentional replacement
is authorized. Clearing has no tool-level undo; copying, sorting, and text
conversion can affect more cells than the starting address suggests. Check
the full intended scope, and do not assume a failed write rolls back earlier changes.

Formatting, protection, and visibility have different effects: hiding formulas
is not encryption, showing rows does not clear filters, and changing a display
format does not convert the underlying data.

Formula relationship inspection is not a complete workbook dependency graph.
Excel's native inspection can omit cross-sheet, external, and dynamic references.
Newer formula and comment capabilities depend on the installed Excel version.

[Range workflow guidance](../reference/range.md) |
[Optional report formatting](../reference/report-formatting.md)

---

## Worksheets (35 operations)

- **Organize sheets:** Create, rename, copy, move, and delete worksheets within a workbook or transfer them between files.
- **Manage presentation:** Set tab colors and visibility, add notes, images, and basic shapes, and organize row/column outlines.
- **Protect editing:** Inspect and configure worksheet protection and permitted interactions.
- **Prepare output:** Configure print areas, repeated titles, margins, headers, scaling, and manual page breaks.

Deleting a worksheet removes its contents and can break dependent references.
Moving a sheet to another file removes the source and saves both workbooks;
closing another session without saving cannot reverse that transfer.
If a cross-file save fails, the error identifies whether the source save was
confirmed. Temporary sessions close without another save; inspect both files
before retrying. These operations do not roll back earlier successful saves.

Protection controls editing, not confidentiality. Some permissions still require
unlocked cells, and automation-only protection is not retained after reopening.
Print layout depends on the printer and scaling. Page setup does not print or
open a modal preview.

[Worksheet workflow guidance](../reference/worksheet.md)

---

## Workbook (27 operations)

- **Inspect and describe:** Read workbook state and manage document properties, protection, and view settings.
- **Use themes and styles:** Inspect or apply Office themes and create, inspect, update, or delete supported custom cell and Table styles.
- **Save and publish:** Change supported file formats, save a requested copy, or export PDF/XPS output.
- **Manage external links:** Inspect or refresh linked workbook sources, or deliberately replace linked formulas with their values.

Changing a theme or shared style can affect existing content throughout the
workbook. Built-in styles are inspectable but not editable through the custom
style lifecycle. Native failures do not promise rollback.

Changing file formats can remove features; saving a macro-enabled workbook as
`.xlsx` removes its VBA content. Breaking an external link is permanent once
saved. Printing and print preview are intentionally excluded from unattended
automation.

[Saving and external-link guidance](../reference/workbook.md) |
[Reusable style guidance](../reference/range.md#reusable-cell-styles)

---

## Named Ranges (Parameters) (6 operations)

- **Drive workbook inputs:** Discover, read, write, and maintain user-defined named ranges used by formulas and queries.

Named ranges refer to cells, not independent stored values. Internal and hidden
Excel names are omitted from discovery. Updating a parameter does not itself
refresh Power Query; refresh the affected load when its results must change.

[Names and data interpretation](../reference/range.md#links-comments-and-names)

---

## Related feature areas

- [Data & analytics](DATA-ANALYTICS.md) - transform data and build analytical summaries
- [Charts & visualization](CHARTS-VISUALS.md) - create visuals from workbook data
- [Automation & advanced](AUTOMATION-ADVANCED.md) - run code and compare assumptions
- [Real Excel automation vs. file-parser libraries](../guides/EXCEL-COM-VS-FILE-PARSERS.md)
