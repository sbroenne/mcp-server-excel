# Slicers

PivotTable slicers filter their connected PivotTables. Table slicers filter one
Excel Table. Filtering source-table rows does not automatically filter a
separate PivotTable cache. Use the matching creation, listing, selection, and
deletion actions for each type.

## Required creation inputs

Every creation needs a session, a unique slicer name, a destination worksheet, and
a cell position for its top-left corner. Names and positions are **not**
generated automatically.

PivotTable slicers additionally need the PivotTable name and field name.
Table slicers need the Table name and column name. Inspect the relevant slicer
list and source fields before creating one. The destination sheet must exist.

These examples assume the session, `SalesPivot`, `SalesData`, and `Analysis` sheet
already exist. Check each result before proceeding.

```mcp
slicer(action: 'create-slicer', session_id: sessionId, pivot_table_name: 'SalesPivot', field_name: 'Region', slicer_name: 'RegionSlicer', destination_sheet: 'Analysis', position: 'E2')
slicer(action: 'set-slicer-selection', session_id: sessionId, slicer_name: 'RegionSlicer', selected_items: '["North"]')
slicer(action: 'create-table-slicer', session_id: sessionId, table_name: 'SalesData', column_name: 'Product', slicer_name: 'ProductSlicer', destination_sheet: 'Analysis', position: 'H2')
slicer(action: 'set-table-slicer-selection', session_id: sessionId, slicer_name: 'ProductSlicer', selected_items: '["Laptop"]')
```

```cli
excelcli -q slicer create-slicer --session $sessionId --pivot-table-name SalesPivot --field-name Region --slicer-name RegionSlicer --destination-sheet Analysis --position E2
excelcli -q slicer set-slicer-selection --session $sessionId --slicer-name RegionSlicer --selected-items '["North"]'
excelcli -q slicer create-table-slicer --session $sessionId --table-name SalesData --column-name Product --slicer-name ProductSlicer --destination-sheet Analysis --position H2
excelcli -q slicer set-table-slicer-selection --session $sessionId --slicer-name ProductSlicer --selected-items '["Laptop"]'
```

## Selection and verification

Selections are JSON-array **text**, such as `'["North","South"]'`, not a native
array argument in MCP. `'[]'` clears the filter. The default replaces the
selection; disabling clear-first adds to it. The implementation compares names
case-insensitively, but use the actual item names returned by Excel.

To switch an existing single-item filter, pass the replacement item to
`set-slicer-selection` or `set-table-slicer-selection`. Leave `clear_first`
omitted or set it to `true` in MCP; omit `--clear-first` or pass
`--clear-first true` in CLI. The new selection replaces the old one, even when
the items have no overlap.

Data Model/OLAP PivotTable slicers use the same PivotTable slicer actions.
Create with the discovered hierarchy name, such as `[Quarters].[Quarter]`.
`availableItems` and `selectedItems` contain the displayed captions. Selection
accepts those captions or MDX unique names; unknown or ambiguous values fail
before changing the filter. An ambiguous-caption error includes the matching MDX
unique names as a JSON array; retry with the intended name from that array.
Adding to an unfiltered slicer validates the requested items but keeps the filter
cleared, so new members introduced by a later Data Model refresh remain visible.

For regular PivotTable and Table slicers, unmatched values are not individually
rejected, and Excel may retain a selection when asked to deselect every item.
Never infer success from the requested values:
read the slicer selection and the filtered Table rows or PivotTable data.
Check combined filters together. Deleting a slicer is not the same operation as
clearing its filter; explicitly clear first if that is the intended result.

Read [PivotTable guidance](pivottable.md) for source refresh and field setup.

## Timelines, shared connections, and layout

`create-timeline` creates a native date control from a PivotTable date field, not
from an ordinary Table column. Its required inputs match `create-slicer`.
Use `get-slicer` for complete native control state: layout/style, all connections,
item selections, and timeline date/view state. `list-slicers` identifies timelines
with `isTimeline` and `timeline`; they do not have ordinary selected-item lists.

`set-timeline-selection` takes MCP `timeline_selection` or CLI
`--timeline-selection`, with nested `startDate` and `endDate` calendar dates.
The inclusive range filters **every connected PivotTable**. Do not use ordinary
item selection for a timeline. `clear-timeline-selection` clears this date
filter; `delete-slicer` removes the visual control, not necessarily its filter.

`connect-pivottable` / `disconnect-pivottable` take MCP `slicer_name` /
`pivot_table_name` or CLI `--slicer-name` / `--pivot-table-name`. Connections
require the existing shared PivotCache; matching field names alone are not
enough. Incompatible caches and Table slicers are rejected without rebuilding
anything. Disconnecting keeps at least one source connection.

`update-slicer` takes MCP `slicer_options` or CLI `--slicer-options`. Nested keys
remain camelCase. Geometry uses points; omitted properties remain unchanged.
Ordinary controls use `columnCount` / `displayHeader`; timelines use `granularity`
and their `showHeader` / `showSelectionLabel` / `showTimeLevel` /
`showHorizontalScrollbar` view settings. Use existing native style names.
Read the resulting state rather than assuming Excel accepted the requested
appearance.
