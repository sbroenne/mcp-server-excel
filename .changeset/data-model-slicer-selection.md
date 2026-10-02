---
"excelmcp": patch
---

**Data Model slicers** (#980): Creating and listing slicers now shows their actual items and selections. Select items by their displayed captions to filter connected PivotTables, replace or add to a selection, or clear the filter; invalid items return an error without changing the selection.

Slicer creation checks cancellation before creating each cache or visual, and connected-PivotTable scans observe cancellation. Cancellation stops further work; it does not undo changes already made.
