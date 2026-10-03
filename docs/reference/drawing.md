# drawing - Server Quirks

Use `drawing` for worksheet images, AutoShapes, text boxes, connectors, safe Forms controls, and sparklines.

Object names are worksheet-local. Call `list-objects` before updates or deletion when the exact name is unknown.

Positions use points, not cells; use the schema/help for styles and placement.

## Drawing layout

`group-objects`, `align-objects`, and `distribute-objects` take `object_names`
(MCP) / `--object-names` (CLI), a JSON array string of distinct top-level names on the
specified worksheet. Grouping and alignment need at least two objects;
distribution needs three. Alignment and equal-gap spacing use the selected
extent, not the worksheet or printed page, and leave unselected objects alone.

`ungroup-object`, `duplicate-object`, and `set-z-order` use `object_name`
(MCP) / `--object-name` (CLI). Grouping and duplication return Excel's actual
new name unless `group_name` / `--group-name` or `new_name` / `--new-name` is
supplied. Supplied names must be unique on the worksheet. Duplication offsets
use `offset_left` / `--offset-left` and `offset_top` / `--offset-top`, in points
from the original position; both default to 10.

These actions return `drawingObjects`, including real positions, stacking
positions and complete group `children`. Ungrouping returns the newly exposed
direct members, not every other worksheet object. Excel can flatten an existing
group when regrouping; inspect the returned members rather than assuming the
old hierarchy survives. `get-object` and `list-objects` also inspect complete
native membership. Stacking positions on top-level objects start at one at the
back; member positions are Excel's native values.

Layout rejects protected drawing objects, charts, ActiveX/OLE and unknown types.
Use the chart tools for chart changes. Duplication also rejects any macro-bound
object or group member, rather than copying a macro assignment.

## Safe Forms controls

- `linked_cell` (MCP) / `--linked-cell` (CLI): CheckBox, DropDown, ListBox, OptionButton, ScrollBar, and Spinner
- `input_range` (MCP) / `--input-range` (CLI): DropDown and ListBox only
- Button, GroupBox, and Label return explicit nulls for both binding properties

ActiveX/OLE controls and macro assignment are intentionally unavailable. Do not try to create them through VBA as a workaround.

## Sparklines

- `source_range` (MCP) / `--source-range` (CLI): data to visualize
- `location_range` (MCP) / `--location-range` (CLI): cells that host the sparklines
- Line sparklines can show markers

```mcp
drawing(action: 'add-shape', session_id: sessionId, sheet_name: 'Dashboard', shape_type: 'RoundedRectangle', name: 'Status', text: 'Ready', fill_color: '#70AD47')
drawing(action: 'add-sparkline', session_id: sessionId, sheet_name: 'Dashboard', source_range: 'B2:E2', location_range: 'F2', sparkline_type: 'Line')
```

```cli
excelcli -q drawing add-shape --session $sessionId --sheet Dashboard --shape-type RoundedRectangle --name Status --text Ready --fill-color '#70AD47'
excelcli -q drawing add-sparkline --session $sessionId --sheet Dashboard --source-range B2:E2 --location-range F2 --sparkline-type Line
```
