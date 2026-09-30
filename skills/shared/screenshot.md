# Screenshots and Visual Verification

Use screenshots when appearance matters, such as chart placement or a report
layout, and an interactive desktop is available. Do not add a screenshot step
to an unrelated read or data-only task. A successful data operation does not
depend on capturing an image.

## Actions

| Action | Framing | Required context |
|--------|---------|--------|
| `capture` | Explicit cell range, default `A1:Z30` | Session and worksheet; optional range and quality |
| `capture-sheet` | Used cells and embedded charts | Session and worksheet; optional quality |

Capture photographs the live Excel window, briefly showing it and bringing it
forward. It requires an unlocked interactive desktop; disconnected Remote
Desktop sessions may prevent capture. Protected sheets and initially hidden
windows are supported. The workbook and clipboard are not modified.

Large ranges are zoomed, captured in passes, and stitched. If the result reports
truncation, use a smaller range rather than claiming the whole area was checked.

| Quality | Format | Scale |
|---------|--------|-------|
| `Medium` (default) | JPEG | 75% |
| `Low` | JPEG | 50% |
| `High` | PNG | 100% |

MCP returns native image content and structured capture metadata. CLI can save
an image directly; use a matching quality and extension:

```cli
excelcli -q screenshot capture --session $sessionId --sheet Sales --range A1:M25 --quality High --output screenshot.png
```

## Layout Checks

For a requested chart, inspect the used range, create or move the chart, and
check returned overlap warnings. A target cell range makes explicit layouts easier;
omitting both it and point coordinates uses supported automatic positioning.

```mcp
chart(action: 'create-from-range', session_id: sessionId, sheet_name: 'Sales',
      source_range_address: 'A1:D20', chart_type: 'ColumnClustered',
      target_range: 'F2:K15')
screenshot(action: 'capture', session_id: sessionId, sheet_name: 'Sales',
           range_address: 'A1:M25')
```

```cli
excelcli -q chart create-from-range --session $sessionId --sheet Sales --source-range-address A1:D20 --chart-type ColumnClustered --target-range F2:K15
excelcli -q screenshot capture --session $sessionId --sheet Sales --range A1:M25 --quality High --output screenshot.png
```

Use field operations to add row/value fields and refresh the PivotTable before
checking its layout. For multiple charts, leave room between them and reposition
with `chart fit-to-range` when needed. Check again after a meaningful layout fix,
not after every routine write.

If capture is unavailable, inspect chart positions and sizes using `chart read`
and explain that visual verification could not be completed. Do not repeatedly
retry an unavailable desktop or prevent an authorized save/close solely because
capture failed.
