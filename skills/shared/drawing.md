# drawing - Server Quirks

Use `drawing` for worksheet images, AutoShapes, text boxes, connectors, safe Forms controls, and sparklines.

## Object lifecycle

| Action | Purpose |
|--------|---------|
| `list-objects` | List drawing objects on one worksheet |
| `get-object` | Read one object by name |
| `add-image` | Embed a local image |
| `add-shape` | Add a geometric, arrow, or flowchart AutoShape |
| `add-text-box` | Add formatted text |
| `add-connector` | Add straight, elbow, or curved connectors |
| `add-form-control` | Add a worksheet Forms control |
| `update-object` | Rename, move, resize, rotate, format, or change bindings |
| `delete-object` | Delete by object name |

Object names are worksheet-local. Call `list-objects` before updates or deletion when the exact name is unknown.

Colors use `#RRGGBB`. Position and size values use points. Placement values are:

- `1`: move and size with cells
- `2`: move but do not size with cells
- `3`: free floating

## Safe Forms controls

Supported controls are Button, CheckBox, DropDown, GroupBox, Label, ListBox, OptionButton, ScrollBar, and Spinner.

- `linked_cell` (MCP) / `--linked-cell` (CLI): CheckBox, DropDown, ListBox, OptionButton, ScrollBar, and Spinner
- `input_range` (MCP) / `--input-range` (CLI): DropDown and ListBox only
- Button, GroupBox, and Label return explicit nulls for both binding properties

ActiveX/OLE controls and macro assignment are intentionally unavailable. Do not try to create them through VBA as a workaround.

## Sparklines

Use `add-sparkline`, `get-sparkline`, `list-sparklines`, `update-sparkline`, and `delete-sparkline`.

- Types: Line, Column, WinLoss
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
