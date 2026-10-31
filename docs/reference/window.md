# Window management

Window changes affect only the selected session's Excel instance. Do not create
another session just to change visibility. Use CLI help or MCP tool descriptions
for current window controls and inputs.

## Worksheet views

Inspect the active sheet, selection, cell, and chart when that context matters.
Context inspection does not show, activate, or select anything, and another
workbook's selection is not a substitute for unavailable context.

Frozen panes retain rows and columns above/left of a boundary; movable splits
are different and replace frozen panes. Set zoom and display choices before a
split when exact pane counts matter, then inspect the resulting view.

### Prepare the view before saving

Use `set-view` to scroll a named sheet and optionally select a cell or continuous
range. The sheet stays active. Omitting selection keeps its existing selection.
The command does not save the workbook; explicitly save after checking the view.

```cli
excelcli -q window set-view --session $sessionId --sheet-name Dashboard --top-left-cell A1 --select A1
excelcli -q window get-view --session $sessionId --sheet-name Dashboard
```

```mcp
window(action: 'set-view', workbook_session_id: sessionId, sheet_name: 'Dashboard', top_left_cell: 'A1', select: 'A1')
window(action: 'get-view', workbook_session_id: sessionId, sheet_name: 'Dashboard')
```

Both actions return `topLeftCell`, `selectionKind`, `selectedRangeAddress`, and
`activeCellAddress`, alongside panes, zoom, and display settings. Addresses are
Excel's normalized addresses. Non-range selections have no range or active-cell
address. `get-view` activates the requested sheet but does not select or scroll it.

Frozen panes and movable splits are retained. The scroll position belongs to the
last pane: the bottom-right pane for splits in both directions, or the bottom/right
pane for a single split. Frozen or hidden rows and columns may make Excel adjust
the requested position; inspect `topLeftCell` for the actual result. Panes share
scrolling in any direction that is not split, following Excel's normal behavior. Neither action
changes the view mode, zoom, or display options, or intentionally shows Excel.

`--top-left-cell` / `top_left_cell` must identify one cell on the named sheet;
`--select` / `select` accepts one cell or a continuous range on that sheet.
References to other sheets or workbooks and multi-area selections are rejected.
If Excel fails after a view change, do not assume the change was undone.
Inspect the view before deciding what to do next.

## Visibility and placement

Preserve existing visibility unless a change is requested or a blocking dialog
is confirmed. Such dialogs automatically reveal Excel so the user can respond.
Keeping a workbook
open means retaining its session, not showing a hidden window. See the shared
[visibility policy](behavioral-rules.md#visibility).

`show` and `bring-to-front` restore a minimized window to its previous normal
or maximized state before bringing it forward. `bring-to-front` leaves a hidden
session hidden and returns guidance to use `show` first. If Windows refuses
foreground activation, `show` still reports the verified visibility and
`isForeground: false`, with guidance to select Excel. `bring-to-front` reports
an error with instructions to select Excel on the Windows taskbar or press
Alt+Tab to switch to Excel. Any visibility or window restoration already applied
remains in effect.
`arrange` verifies that Excel applied the requested preset and returns
`isForeground`. If Windows refuses foreground activation, the verified
placement still succeeds and the result explains how to select Excel. If Excel
does not apply the requested placement, the operation reports an error; any
window changes already applied remain in effect.
Standalone `show` and `hide` also work while Excel is waiting on a dialog,
without waiting behind the blocked command. Neither answers nor dismisses it.
Use `window(action: 'hide', workbook_session_id: sessionId)` or
`excelcli window hide --session $sessionId` to hide explicitly.
An unsupported arrange `preset` (MCP) / `--preset` (CLI) is rejected before
changing visibility, window state, or bounds. This does not promise rollback
if Excel itself fails while applying a supported preset.

Arranging or restoring a normal/maximized window can make it visible. Layout
uses the monitor containing Excel, and positioning uses points rather than
pixel or cell counts. Do not assume those changes preserve hidden mode.

For requested side-by-side work:

```mcp
window(action: 'show', workbook_session_id: sessionId)
window(action: 'arrange', workbook_session_id: sessionId, preset: 'right-half')
```

```cli
excelcli -q window show --session $sessionId
excelcli -q window arrange --session $sessionId --preset right-half
```

Visible work needs no extra charts or formatting. Optional status text should
be cleared after success or failure. Do not tell a user to inspect a hidden
window. [Screenshots](screenshot.md) can bring Excel forward and need an
interactive desktop.

Close only when authorized and active work has finished. Keep intended edits
through explicit saving; see [session recovery](behavioral-rules.md#sessions-and-failures).
