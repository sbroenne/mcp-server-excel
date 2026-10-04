# Worksheets

Inspect existing names, contents, and dependencies before creating, moving,
or removing sheets. Deleting a worksheet removes all its contents and can
break references. There is no tool-level undo.

Use CLI help or MCP tool descriptions for current worksheet controls and inputs.

The server checks Microsoft's documented naming rules before creating, copying,
or renaming a worksheet: names cannot be blank, exceed 31 characters, contain
`/ \ ? * : [ ]`, begin or end with an apostrophe, or be `History`.
Non-English names and apostrophes inside a name are allowed. See Microsoft's
[worksheet naming guidance](https://support.microsoft.com/en-us/excel/rename-a-worksheet).
Names are not trimmed: leading or trailing spaces may be part of a name, but a
name made entirely of whitespace is rejected as blank.
The server checks these documented constraints before changing the workbook;
Excel also validates the name when it is assigned. If Excel rejects a name
after a create or copy has added the sheet, the error identifies that sheet.
It remains in the workbook for the caller to inspect and remove if appropriate.

## Cross-file operations

Use the user's intended source and destination files. Cross-file transfer owns
opening, saving, and closing those files; do not apply it to workbooks already
open in unrelated sessions.

Copying keeps the source sheet. Moving removes it and saves both workbooks.
Closing another session without saving cannot reverse a saved transfer.
Failures do not promise rollback across files: inspect both before retrying.

## Styling and outlines

Worksheet presentation includes tab colors, visibility, legacy notes, images,
shapes, page layout, and outlines. Legacy notes are different from modern
threaded comments.

Page setup prepares output but does not print or open preview. Decide whether
fixed scaling or fitting to pages serves the report; printer support and
scaling affect automatic pagination. Manual-break replacement affects the
whole sheet, including breaks outside the current print area.

Inspect the print area, repeated headers, margins, and page breaks before
[PDF/XPS export](workbook.md#save-and-publish). Native settings can partly change
before failure, so read back the result rather than assuming rollback.

Outline changes affect complete rows or columns. Inspect existing groups and
summary positions first. Removing one grouping level is different from clearing
all outlines; expand/collapse when only the visible detail level should change.

## Protection permissions

Changing worksheet protection replaces its permission configuration; omitted
permissions need not preserve the previous choices. Inspect the full current
state before replacing it.

Filtering permission allows changes to an existing filter, not necessarily
creating a new one. Sorting and deletion can still require unlocked cells.
Automation-only protection is runtime-only and is not retained after reopening.

Worksheet protection and hidden formulas control editing or display, not
encryption or confidentiality. Do not bypass security or promise that a failed
change restored the previous state.

## Workbook themes

A theme change can affect colors and fonts throughout the workbook, including
styles and charts. Fixed RGB colors remain fixed. Choose an existing Office
theme file for the requested appearance rather than making unrelated changes.

Inspect the actual theme definitions. Empty script-font definitions do not mean
fallback fonts were applied, and Excel cannot expose every theme name or original
file path. Read back changes before retrying after failure; saving remains explicit.
