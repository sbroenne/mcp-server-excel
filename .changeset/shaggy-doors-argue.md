---
"excelmcp": patch
---

**Screenshots now photograph the live Excel window** (#777) — `screenshot capture` and `capture-sheet` used to render the image inside Excel by inserting a temporary chart into the worksheet. On a protected sheet Excel refuses that insert, so capture failed with a bare `COMException 0x800A03EC`. Capture now takes a real picture of the Excel window instead.

What this changes for you:

- Protected sheets can be captured.
- Your clipboard is no longer overwritten, and the workbook is never modified or dirtied by taking a screenshot.
- Captures are faster, since the old chart create/paste/export retry ladder is gone.
- Ranges larger than the Excel window are zoomed to fit and, if still too large, captured in several passes and stitched together; extremely large ranges are truncated to their top-left portion and the result message says so.

Capture now requires an interactive desktop session — it will fail on a locked desktop or a disconnected Remote Desktop session.
