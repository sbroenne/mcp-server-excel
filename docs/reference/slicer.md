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
