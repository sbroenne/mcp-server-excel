# Slicers

A PivotTable slicer filters its connected PivotTables. A Table slicer filters
one worksheet Table, not a separate PivotTable cache. A dashboard control does
not automatically filter every chart.

Current commands and inputs come from CLI help or MCP tool descriptions.

## Plan the control {#required-creation-inputs}

Inspect source fields, existing control names, and the destination sheet before
creating a control. Choose its location as part of the requested layout rather
than assuming Excel will select a suitable name or position.

Use a date timeline for a PivotTable date field, not an ordinary Table column.
A live PivotChart follows the connected PivotTable; a regular chart may not.

Creation and selection results, and listed slicers, identify the source with
`sourceType` (`"PivotTable"` or `"Table"`). Returned `position` is Excel's
actual `Shape.TopLeftCell`, not an echo of the requested position. Excel rounds
shape coordinates, so a shape aligned with a cell boundary can report the
preceding cell. Verify placement using the native coordinates as well as the
reported anchor when testing exact positioning.

## Selection and verification

Choose whether to replace the existing selection or add to it. Use actual
returned items instead of guessing labels. Data Model captions can be ambiguous;
use the discovered unique item name when needed. Adding to an unfiltered model
slicer keeps it unfiltered, allowing new members after refresh.

For regular and Table slicers, the resulting selection can differ from the
request when items do not match or Excel will not deselect every item. Read both
the selected items and the filtered rows or PivotTable results. Check combined
filters together.

Removing a slicer is not the same as clearing its filter. Clear the filter
first when that is the intended outcome, then inspect the resulting data.

See [PivotTable guidance](pivottable.md) for source refresh and field setup.

## Timelines, shared connections, and layout

A timeline filters an inclusive calendar-date range on every connected
PivotTable. It uses date selection, not ordinary item selection. Clearing its
date filter does not clear unrelated filters.

Shared controls can connect to PivotTables using the same existing compatible
cache. Matching field names alone are not enough; connections do not rebuild
incompatible caches. Table slicers cannot connect to PivotTables, and a
PivotTable control must keep at least one source connection.

Inspect actual connections before promising a shared dashboard filter.
After a layout/style change, read the control's resulting appearance and bounds
instead of assuming that every requested setting was accepted by Excel.

## Custom slicer styles

Use `slicer list-styles` to discover built-in and custom styles available in the
selected workbook. `create-style` copies an existing slicer style into a new
workbook-owned style; it does not apply the copy automatically. Existing names
are rejected. `apply-style` applies a named style to one ordinary PivotTable,
Data Model, or Table slicer without changing its column count, size, position,
caption, connections, or filters.

For example, in PowerShell:

```powershell
excelcli slicer list-styles --session $session
$options = '{"selectedItemWithData":{"fillColor":"#E8F0FE","borderColor":"#336699","fontColor":"#123456","bold":true},"unselectedItemWithData":{"fillColor":"#FFFFFF","borderColor":"#808080","bold":false},"unselectedItemWithNoData":{"fillColor":"#E0E0E0"}}'
excelcli slicer create-style --session $session --style-name Brand --source-style-name SlicerStyleLight1 --slicer-style-options $options
excelcli slicer apply-style --session $session --slicer-name Regions --style-name Brand
```

In MCP, call `slicer` with `action: "create-style"`, `workbook_session_id`,
`style_name`, `source_style_name`, and a JSON object in `slicer_style_options`.
Applying it uses
`action: "apply-style"`, `slicer_name`, and `style_name` in the same session.
Nested options keep their camelCase keys, as in the CLI example.

Each supplied element accepts `fillColor`, `borderColor`, `fontColor` in
`#RRGGBB` form and `bold`. Border colour sets four solid, thin outer borders.
Omitted elements and fields retain their source formatting; `bold: false`
explicitly removes bold. An empty options object makes an unchanged copy.
The options cover the frame, header, selected/unselected buttons with or
without data, and the four corresponding hovered states.

Inspect a full definition with `workbook get-table-style`. For later edits or
deletion, use `workbook update-table-style` or `workbook delete-table-style`;
these can affect every control using that style. A failure after copying can
leave a partially formatted style in the workbook. Inspect the reported name
before deciding whether to edit or delete it; failures do not undo changes.

These three slicer-style commands do not handle timelines. For a timeline,
use the existing workbook style commands and `slicer update-slicer` with the
`style` field inside `slicerOptions` (CLI: `--slicer-options`; MCP:
`slicer_options`).
