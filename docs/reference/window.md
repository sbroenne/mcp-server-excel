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

## Visibility and placement

Preserve existing visibility unless a change is requested. Keeping a workbook
open means retaining its session, not showing a hidden window. See the shared
[visibility policy](behavioral-rules.md#visibility).

`show` and `bring-to-front` restore a minimized window to its previous normal
or maximized state before bringing it forward. `bring-to-front` leaves a hidden
session hidden and returns guidance to use `show` first. If Windows refuses
foreground activation, the operation reports an error; any visibility or
window restoration already applied remains in effect.
An unsupported arrange `preset` (MCP) / `--preset` (CLI) is rejected before
changing visibility, window state, or bounds. This does not promise rollback
if Excel itself fails while applying a supported preset.

Arranging or restoring a normal/maximized window can make it visible. Layout
uses the monitor containing Excel, and positioning uses points rather than
pixel or cell counts. Do not assume those changes preserve hidden mode.

For requested side-by-side work:

```mcp
window(action: 'show', session_id: sessionId)
window(action: 'arrange', session_id: sessionId, preset: 'right-half')
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
